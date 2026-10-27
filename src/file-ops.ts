/**
 * Firm file tools: move bytes between Outlook mail and OneDrive or SharePoint
 * entirely inside the server, Graph to Graph. No tool here ever returns file
 * bytes (base64 or otherwise) to the caller; each returns the ids, names, sizes,
 * and links of what it wrote.
 *
 * Tools:
 *   save-attachment-to-drive         mail attachment -> drive folder
 *   attach-drive-file-to-draft       drive file -> draft message attachment
 *   attach-mail-attachment-to-draft  mail attachment -> draft message attachment
 *   find-matter-folder               search the Matters tree (three levels under grouping folders)
 *   file-email-to-matter             message (.eml) and its attachments -> matter folder
 *
 * Every call writes one log line: { ts, tool, source, destination, size }.
 */
import { z } from 'zod';
import logger from './logger.js';
import { getUserIdentityForAudit } from './audit-log.js';
import { getGraphBaseUrl } from './cloud-config.js';
import type GraphClient from './graph-client.js';
import type { UtilityTool, UtilityToolContext } from './graph-tools.js';
import {
  fetchWithResilience,
  getSharedBreaker,
  loadResilienceConfig,
} from './lib/graph-resilience.js';
import { getRequestTokens } from './request-context.js';

const MB = 1024 * 1024;
/** Largest file any tool here will move (the Graph upload session ceiling for attachments). */
export const FILE_OPS_MAX_BYTES = 150 * MB;
/** Drive uploads below this size use a single PUT; at or above it, an upload session. */
export const DRIVE_SIMPLE_UPLOAD_MAX_BYTES = 4 * MB;
/** Mail attachments below this size are POSTed inline; at or above it, an upload session. */
export const MAIL_INLINE_ATTACHMENT_MAX_BYTES = 3 * MB;
/**
 * Upload session chunk size. Drive sessions require a multiple of 320 KiB; Outlook
 * attachment sessions require each range to be under 4 MB. 10 x 320 KiB meets both.
 */
export const UPLOAD_CHUNK_BYTES = 320 * 1024 * 10;

export const DEFAULT_MATTERS_DRIVE_ID =
  'b!m2Rv4o9xZE-3OoioO8WFPI4O6eEnFgZIgL385_dGYq9KvqgnJ-TBTqHv8y1m3SlH';
export const DEFAULT_MATTERS_ROOT_PATH = 'GC Associates/Matters';

const CONFLICT_BEHAVIORS = ['rename', 'fail', 'replace'] as const;
export type ConflictBehavior = (typeof CONFLICT_BEHAVIORS)[number];

interface ToolResult {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
  [key: string]: unknown;
}

/** A failure the tool reports to the caller as a plain, clear message. */
export class FileOpsError extends Error {}

export function ok(payload: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload) }] };
}

export function fail(message: string): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify({ error: message }) }], isError: true };
}

export function logOp(
  tool: string,
  source: string,
  destination: string,
  size: number | null
): void {
  logger.info(
    `[file-ops] ${JSON.stringify({ ts: new Date().toISOString(), tool, source, destination, size })}`
  );
}

/**
 * Strip the characters SharePoint and OneDrive refuse in a file name
 * (: # % < > " | ? * plus the path separators / and \ and control characters),
 * collapse runs of whitespace, and drop leading and trailing spaces and dots.
 */
export function sharePointSafeName(name: string): string {
  const cleaned = name
    .split('')
    .filter((ch) => {
      const code = ch.charCodeAt(0);
      if (code < 32 || code === 127) return false;
      return !':#%<>"|?*/\\'.includes(ch);
    })
    .join('')
    .replace(/\s+/g, ' ')
    .replace(/^[\s.]+|[\s.]+$/g, '');
  return cleaned;
}

export function enc(id: string): string {
  return encodeURIComponent(id);
}

function encodePath(p: string): string {
  return p
    .split('/')
    .filter((s) => s.length > 0)
    .map((s) => encodeURIComponent(s))
    .join('/');
}

export function statusOf(error: unknown): number | undefined {
  const e = error as { httpStatus?: unknown; message?: unknown };
  if (typeof e?.httpStatus === 'number') return e.httpStatus;
  const m = typeof e?.message === 'string' ? e.message.match(/API error: (\d{3})\b/) : null;
  return m ? Number(m[1]) : undefined;
}

/** Turn a Graph failure into a message that names the thing that was not found or refused. */
export function describe(error: unknown, what: string): string {
  if (error instanceof FileOpsError) return error.message;
  const status = statusOf(error);
  const raw = (error as Error)?.message ?? String(error);
  if (status === 404) return `${what} was not found (Graph 404). Check the id. Detail: ${raw}`;
  if (status === 403) return `Access to ${what} was refused (Graph 403). Detail: ${raw}`;
  if (status === 401)
    return `The access token was rejected (Graph 401). Sign in again. Detail: ${raw}`;
  return raw;
}

/**
 * Resolve the Graph token for this call the same way download-bytes does: the
 * request-scoped bearer token in HTTP/OAuth mode, the selected cached account in
 * stdio mode. A missing token is refused before any Graph call.
 */
export async function resolveToken(
  ctx: UtilityToolContext,
  account: string | undefined
): Promise<string> {
  const { authManager } = ctx;
  const contextToken = getRequestTokens()?.accessToken;
  try {
    if (account && authManager && (contextToken || authManager.isOAuthModeEnabled())) {
      const bearer = contextToken ?? (await authManager.getToken().catch(() => null)) ?? undefined;
      const identity = getUserIdentityForAudit(bearer);
      if (!identity || identity.toLowerCase() !== account.toLowerCase()) {
        throw new FileOpsError(
          `The 'account' parameter is not supported in HTTP/OAuth mode: every request uses the identity of the connecting client's bearer token${identity ? ` ('${identity}')` : ''}.`
        );
      }
    }
    let token: string | null | undefined;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      token = await authManager.getTokenForAccount(account);
    } else {
      token = contextToken ?? (authManager ? await authManager.getToken() : undefined);
    }
    if (!token) throw new FileOpsError('missing');
    return token;
  } catch (error) {
    if (error instanceof FileOpsError && error.message !== 'missing') throw error;
    throw new FileOpsError(
      'No Microsoft Graph access token is available for this request. Sign in to the connector again, then retry.'
    );
  }
}

export async function graphJson<T>(
  graph: GraphClient,
  token: string,
  endpoint: string,
  options: {
    method?: string;
    body?: unknown;
    headers?: Record<string, string>;
    quietStatuses?: number[];
  } = {}
): Promise<T> {
  const body =
    options.body === undefined
      ? undefined
      : typeof options.body === 'string' || Buffer.isBuffer(options.body)
        ? (options.body as string | Buffer)
        : JSON.stringify(options.body);
  return (await graph.makeRequest(endpoint, {
    accessToken: token,
    method: options.method ?? 'GET',
    ...(body !== undefined ? { body } : {}),
    ...(options.headers ? { headers: options.headers } : {}),
    ...(options.quietStatuses ? { quietStatuses: options.quietStatuses } : {}),
  })) as T;
}

/** Read Graph byte content into memory, refusing past FILE_OPS_MAX_BYTES. */
export async function graphBytes(
  graph: GraphClient,
  token: string,
  endpoint: string,
  maxBytes: number = FILE_OPS_MAX_BYTES
): Promise<Buffer> {
  const res = await graph.downloadStream(endpoint, { accessToken: token });
  if (res.contentLength !== null && res.contentLength > maxBytes) {
    await res.body.cancel().catch(() => undefined);
    throw new FileOpsError(tooLarge(res.contentLength, maxBytes));
  }
  const reader = res.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      await reader.cancel().catch(() => undefined);
      throw new FileOpsError(tooLarge(total, maxBytes));
    }
    chunks.push(value);
  }
  return Buffer.concat(chunks);
}

export function tooLarge(size: number, maxBytes: number = FILE_OPS_MAX_BYTES): string {
  return `File is ${size} bytes, over the ${Math.round(maxBytes / MB)} MB (${maxBytes} bytes) limit for this tool. Move it in the browser instead.`;
}

/** PUT the buffer to a pre-authenticated upload session URL in ranges. Returns the last response. */
async function uploadInRanges(
  uploadUrl: string,
  bytes: Buffer
): Promise<{ status: number; headers: Headers; text: string }> {
  const config = loadResilienceConfig();
  const breaker = getSharedBreaker();
  let last: { status: number; headers: Headers; text: string } | undefined;
  for (let start = 0; start < bytes.length; start += UPLOAD_CHUNK_BYTES) {
    const end = Math.min(start + UPLOAD_CHUNK_BYTES, bytes.length);
    const chunk = bytes.subarray(start, end);
    // The session URL carries its own authorization; Graph rejects a bearer header here.
    const response = await fetchWithResilience(
      uploadUrl,
      {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/octet-stream',
          'Content-Length': String(chunk.length),
          'Content-Range': `bytes ${start}-${end - 1}/${bytes.length}`,
        },
        body: chunk as unknown as string,
      },
      config,
      breaker
    );
    const text = await response.text();
    if (!response.ok) {
      await fetch(uploadUrl, { method: 'DELETE' }).catch(() => undefined);
      throw new FileOpsError(
        `Upload session rejected bytes ${start}-${end - 1} of ${bytes.length}: ${response.status} ${text.slice(0, 500)}`
      );
    }
    last = { status: response.status, headers: response.headers, text };
  }
  if (!last) throw new FileOpsError('Nothing to upload: the file is empty.');
  return last;
}

// ---------------------------------------------------------------------------
// Sources

export interface SourceFile {
  bytes: Buffer;
  name: string;
  contentType: string;
  label: string;
}

interface AttachmentMeta {
  '@odata.type'?: string;
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
}

export async function readMailAttachment(
  graph: GraphClient,
  token: string,
  messageId: string,
  attachmentId: string,
  maxBytes: number = FILE_OPS_MAX_BYTES
): Promise<SourceFile> {
  const base = `/me/messages/${enc(messageId)}/attachments/${enc(attachmentId)}`;
  const label = `message:${messageId}/attachment:${attachmentId}`;
  let meta: AttachmentMeta;
  try {
    // $select keeps contentBytes out of the metadata read.
    meta = await graphJson<AttachmentMeta>(
      graph,
      token,
      `${base}?$select=id,name,contentType,size`
    );
  } catch (error) {
    throw new FileOpsError(describe(error, `Attachment ${attachmentId} on message ${messageId}`));
  }
  const kind = (meta['@odata.type'] ?? '').toLowerCase();
  if (kind.endsWith('referenceattachment')) {
    throw new FileOpsError(
      `Attachment "${meta.name ?? attachmentId}" is a cloud link (reference attachment), not a file. Open the linked file instead.`
    );
  }
  if (typeof meta.size === 'number' && meta.size > maxBytes) {
    throw new FileOpsError(tooLarge(meta.size, maxBytes));
  }
  let bytes: Buffer;
  try {
    bytes = await graphBytes(graph, token, `${base}/$value`, maxBytes);
  } catch (error) {
    throw new FileOpsError(describe(error, `Attachment ${attachmentId} content`));
  }
  let name = meta.name || 'attachment';
  if (kind.endsWith('itemattachment') && !/\.[a-z0-9]{2,5}$/i.test(name)) name = `${name}.eml`;
  return {
    bytes,
    name,
    contentType: meta.contentType || 'application/octet-stream',
    label,
  };
}

interface DriveItemMeta {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  file?: { mimeType?: string };
  folder?: unknown;
}

export async function readDriveFile(
  graph: GraphClient,
  token: string,
  driveId: string,
  itemId: string,
  maxBytes: number = FILE_OPS_MAX_BYTES
): Promise<SourceFile> {
  const base = `/drives/${enc(driveId)}/items/${enc(itemId)}`;
  let meta: DriveItemMeta;
  try {
    meta = await graphJson<DriveItemMeta>(
      graph,
      token,
      `${base}?$select=id,name,size,webUrl,file,folder`
    );
  } catch (error) {
    throw new FileOpsError(describe(error, `Drive item ${itemId} in drive ${driveId}`));
  }
  if (meta.folder !== undefined || meta.file === undefined) {
    throw new FileOpsError(`Drive item "${meta.name ?? itemId}" is a folder, not a file.`);
  }
  if (typeof meta.size === 'number' && meta.size > maxBytes) {
    throw new FileOpsError(tooLarge(meta.size, maxBytes));
  }
  let bytes: Buffer;
  try {
    bytes = await graphBytes(graph, token, `${base}/content`, maxBytes);
  } catch (error) {
    throw new FileOpsError(describe(error, `Drive item ${itemId} content`));
  }
  return {
    bytes,
    name: meta.name || 'file',
    contentType: meta.file?.mimeType || 'application/octet-stream',
    label: `drive:${driveId}/item:${itemId}`,
  };
}

// ---------------------------------------------------------------------------
// Destinations

interface DraftMeta {
  id?: string;
  isDraft?: boolean;
  subject?: string;
  webLink?: string;
}

async function requireDraft(
  graph: GraphClient,
  token: string,
  draftId: string
): Promise<DraftMeta> {
  let draft: DraftMeta;
  try {
    draft = await graphJson<DraftMeta>(
      graph,
      token,
      `/me/messages/${enc(draftId)}?$select=id,isDraft,subject,webLink`
    );
  } catch (error) {
    throw new FileOpsError(describe(error, `Draft message ${draftId}`));
  }
  if (draft.isDraft !== true) {
    throw new FileOpsError(
      `Message ${draftId} ("${draft.subject ?? ''}") is not a draft. Attachments are added to drafts only.`
    );
  }
  return draft;
}

function attachmentIdFromLocation(location: string | null): string | undefined {
  if (!location) return undefined;
  const m = location.match(/Attachments\('([^']+)'\)/i);
  return m ? decodeURIComponent(m[1]) : undefined;
}

async function attachToDraft(
  graph: GraphClient,
  token: string,
  draftId: string,
  file: SourceFile,
  attachmentName: string | undefined
): Promise<{ attachmentId: string; name: string; size: number; method: string }> {
  const name = (attachmentName && attachmentName.trim()) || file.name;
  const size = file.bytes.length;
  const base = `/me/messages/${enc(draftId)}/attachments`;
  if (size < MAIL_INLINE_ATTACHMENT_MAX_BYTES) {
    const created = await graphJson<{ id?: string; name?: string; size?: number }>(
      graph,
      token,
      base,
      {
        method: 'POST',
        body: {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name,
          contentType: file.contentType,
          contentBytes: file.bytes.toString('base64'),
        },
      }
    );
    if (!created?.id) throw new FileOpsError('Graph did not return an attachment id.');
    return { attachmentId: created.id, name: created.name ?? name, size, method: 'inline' };
  }
  const session = await graphJson<{ uploadUrl?: string }>(
    graph,
    token,
    `${base}/createUploadSession`,
    {
      method: 'POST',
      body: {
        AttachmentItem: { attachmentType: 'file', name, size, contentType: file.contentType },
      },
    }
  );
  if (!session?.uploadUrl) throw new FileOpsError('Graph did not return an upload session URL.');
  const last = await uploadInRanges(session.uploadUrl, file.bytes);
  let attachmentId = attachmentIdFromLocation(last.headers.get('location'));
  if (!attachmentId) {
    // Fall back to reading the draft's attachment list (metadata only).
    const list = await graphJson<{ value?: Array<{ id?: string; name?: string; size?: number }> }>(
      graph,
      token,
      `${base}?$select=id,name,size`
    );
    const match = (list.value ?? []).filter((a) => a.name === name).pop();
    attachmentId = match?.id;
  }
  if (!attachmentId) {
    throw new FileOpsError(
      `Upload finished but the new attachment id could not be read. Check the draft's attachment list for "${name}".`
    );
  }
  return { attachmentId, name, size, method: 'uploadSession' };
}

export interface DriveWriteResult {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
}

export async function writeToDrive(
  graph: GraphClient,
  token: string,
  driveId: string,
  folder: { itemId?: string; path?: string },
  fileName: string,
  bytes: Buffer,
  conflictBehavior: ConflictBehavior
): Promise<DriveWriteResult> {
  const parent = folder.itemId
    ? `/drives/${enc(driveId)}/items/${enc(folder.itemId)}:/${enc(fileName)}:`
    : `/drives/${enc(driveId)}/root:/${encodePath(folder.path ?? '')}/${enc(fileName)}:`;
  if (bytes.length < DRIVE_SIMPLE_UPLOAD_MAX_BYTES) {
    return await graphJson<DriveWriteResult>(
      graph,
      token,
      `${parent}/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`,
      {
        method: 'PUT',
        body: bytes,
        headers: { 'Content-Type': 'application/octet-stream' },
      }
    );
  }
  const session = await graphJson<{ uploadUrl?: string }>(
    graph,
    token,
    `${parent}/createUploadSession`,
    {
      method: 'POST',
      body: { item: { '@microsoft.graph.conflictBehavior': conflictBehavior, name: fileName } },
    }
  );
  if (!session?.uploadUrl) throw new FileOpsError('Graph did not return an upload session URL.');
  const last = await uploadInRanges(session.uploadUrl, bytes);
  try {
    return JSON.parse(last.text) as DriveWriteResult;
  } catch {
    throw new FileOpsError('Upload finished but Graph did not return the new drive item.');
  }
}

// ---------------------------------------------------------------------------
// Matter folder search

interface FolderChild {
  id?: string;
  name?: string;
  webUrl?: string;
  folder?: unknown;
}

const rootCache = new Map<string, { id: string; webUrl?: string }>();

/** Test hook: forget the resolved Matters root. */
export function __resetMattersRootCacheForTests(): void {
  rootCache.clear();
}

export function mattersRoot(): { driveId: string; rootPath: string } {
  const driveId = process.env.MATTERS_DRIVE_ID?.trim() || DEFAULT_MATTERS_DRIVE_ID;
  const rootPath = (process.env.MATTERS_ROOT_PATH?.trim() || DEFAULT_MATTERS_ROOT_PATH).replace(
    /^\/+|\/+$/g,
    ''
  );
  return { driveId, rootPath };
}

async function listChildFolders(
  graph: GraphClient,
  token: string,
  driveId: string,
  itemId: string
): Promise<FolderChild[]> {
  const out: FolderChild[] = [];
  let endpoint: string | undefined =
    `/drives/${enc(driveId)}/items/${enc(itemId)}/children?$select=id,name,webUrl,folder&$top=999`;
  const graphBase = `${getGraphBaseUrl()}/v1.0`;
  while (endpoint) {
    const page: { value?: FolderChild[]; '@odata.nextLink'?: string } = await graphJson(
      graph,
      token,
      endpoint
    );
    for (const child of page.value ?? []) if (child.folder !== undefined) out.push(child);
    const next = page['@odata.nextLink'];
    endpoint = next
      ? next.startsWith(graphBase)
        ? next.slice(graphBase.length)
        : next.replace(/^https:\/\/[^/]+\/v1\.0/, '')
      : undefined;
  }
  return out;
}

async function mapLimit<T, R>(
  items: T[],
  limit: number,
  fn: (item: T) => Promise<R>
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i]);
    }
  });
  await Promise.all(workers);
  return results;
}

export const DEFAULT_MATTERS_GROUPING_FOLDERS = 'Other,Customers,Litigation,Insurance Claims';

/**
 * Top-level Matters folders that group clients rather than being clients themselves
 * (Matters/Other/Bodec/NBU Public Information Request). The search descends a third
 * level under these. MATTERS_GROUPING_FOLDERS overrides the list (comma separated).
 */
export function mattersGroupingFolders(): Set<string> {
  const raw = process.env.MATTERS_GROUPING_FOLDERS ?? DEFAULT_MATTERS_GROUPING_FOLDERS;
  return new Set(
    raw
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .filter((s) => s.length > 0)
  );
}

export interface MatterChild {
  name: string;
  itemId: string;
  path: string;
}

export interface MatterMatch {
  /** Grouping folder the client sits under (e.g. "Other"), or null. */
  group: string | null;
  client: string;
  /** Matter folder below the client ("A/B" when two levels down), or null at client level. */
  matter: string | null;
  itemId: string;
  driveId: string;
  /** Path from the drive root, usable as save-attachment-to-drive folderPath. */
  path: string;
  /** Path from the Matters root (e.g. "Other/Bodec/NBU Public Information Request"). */
  mattersPath: string;
  /** 1 = directly under Matters, 2 = one below, 3 = two below. */
  depth: number;
  webUrl: string | null;
  exact: boolean;
  /** Child folders of an exact match, so a sub-folder can be picked without a second search. */
  children?: MatterChild[];
}

interface TreeNode {
  segments: FolderChild[];
}

function describeNode(
  node: TreeNode,
  grouping: Set<string>,
  driveId: string,
  rootPath: string
): Omit<MatterMatch, 'exact' | 'children'> {
  const names = node.segments.map((s) => s.name ?? '');
  const leaf = node.segments[node.segments.length - 1];
  const grouped = names.length >= 2 && grouping.has(names[0].toLowerCase());
  const group = grouped ? names[0] : null;
  const rest = grouped ? names.slice(1) : names;
  const mattersPath = names.join('/');
  return {
    group,
    client: rest[0],
    matter: rest.length > 1 ? rest.slice(1).join('/') : null,
    itemId: leaf.id as string,
    driveId,
    path: `${rootPath}/${mattersPath}`,
    mattersPath,
    depth: names.length,
    webUrl: leaf.webUrl ?? null,
  };
}

// ---------------------------------------------------------------------------
// Filing an email to a matter folder

export const EML_NAME_MAX_CHARS = 120;
export const EMAILS_FOLDER_NAME = 'Emails';
export const DEFAULT_FIRM_TIME_ZONE = 'America/Denver';

/** Remove leading reply and forward prefixes (Re:, RE:, Fw:, FWD:, Re[2]:), repeatedly. */
export function stripReplyPrefixes(subject: string): string {
  let out = subject;
  const prefix = /^\s*(re|fw|fwd)\s*(\[\d+\])?\s*:\s*/i;
  while (prefix.test(out)) out = out.replace(prefix, '');
  return out.trim();
}

/** yyyy-mm-dd of an ISO timestamp in the firm's time zone (FIRM_TIME_ZONE, default Denver). */
export function firmDate(iso: string | undefined, now: Date = new Date()): string {
  const d = iso ? new Date(iso) : now;
  const when = Number.isNaN(d.getTime()) ? now : d;
  const timeZone = process.env.FIRM_TIME_ZONE?.trim() || DEFAULT_FIRM_TIME_ZONE;
  // en-CA formats as yyyy-mm-dd.
  return new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(when);
}

/**
 * "<yyyy-mm-dd> <subject>.eml": reply and forward prefixes stripped, SharePoint-illegal
 * characters removed, the whole name capped at 120 characters.
 */
export function emlFileName(subject: string | undefined, date: string): string {
  const cleanSubject = sharePointSafeName(stripReplyPrefixes(subject ?? '')) || '(no subject)';
  const ext = '.eml';
  const stem = sharePointSafeName(`${date} ${cleanSubject}`)
    .slice(0, EML_NAME_MAX_CHARS - ext.length)
    .replace(/[\s.]+$/g, '');
  return `${stem}${ext}`;
}

/** Find the named child folder, creating it when missing. Refuses a file of that name. */
async function ensureChildFolder(
  graph: GraphClient,
  token: string,
  driveId: string,
  parentItemId: string,
  name: string
): Promise<{ id: string; webUrl?: string; created: boolean }> {
  const lookup = `/drives/${enc(driveId)}/items/${enc(parentItemId)}:/${enc(name)}?$select=id,name,webUrl,folder,file`;
  const existing = async (): Promise<{ id: string; webUrl?: string } | undefined> => {
    try {
      // A missing folder answers 404; that is the normal path, not an error.
      const item = await graphJson<DriveItemMeta>(graph, token, lookup, { quietStatuses: [404] });
      if (item.folder === undefined) {
        throw new FileOpsError(
          `"${name}" in folder ${parentItemId} is a file, not a folder. Nothing was written.`
        );
      }
      return { id: item.id as string, webUrl: item.webUrl };
    } catch (error) {
      if (statusOf(error) === 404) return undefined;
      throw error;
    }
  };
  const found = await existing();
  if (found) return { ...found, created: false };
  try {
    const made = await graphJson<DriveItemMeta>(
      graph,
      token,
      `/drives/${enc(driveId)}/items/${enc(parentItemId)}/children`,
      {
        method: 'POST',
        body: { name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' },
      }
    );
    return { id: made.id as string, webUrl: made.webUrl, created: true };
  } catch (error) {
    // Created by a parallel call between the lookup and the create.
    if (statusOf(error) === 409) {
      const raced = await existing();
      if (raced) return { ...raced, created: false };
    }
    throw error;
  }
}

interface MessageAttachmentRow {
  '@odata.type'?: string;
  id?: string;
  name?: string;
  size?: number;
  isInline?: boolean;
}

async function listMessageAttachments(
  graph: GraphClient,
  token: string,
  messageId: string
): Promise<MessageAttachmentRow[]> {
  const out: MessageAttachmentRow[] = [];
  const graphBase = `${getGraphBaseUrl()}/v1.0`;
  let endpoint: string | undefined =
    `/me/messages/${enc(messageId)}/attachments?$select=id,name,contentType,size,isInline`;
  while (endpoint) {
    const page: { value?: MessageAttachmentRow[]; '@odata.nextLink'?: string } = await graphJson(
      graph,
      token,
      endpoint
    );
    out.push(...(page.value ?? []));
    const next = page['@odata.nextLink'];
    endpoint = next
      ? next.startsWith(graphBase)
        ? next.slice(graphBase.length)
        : next.replace(/^https:\/\/[^/]+\/v1\.0/, '')
      : undefined;
  }
  return out;
}

interface FiledItem {
  kind: 'eml' | 'attachment';
  itemId?: string;
  name: string;
  size: number;
  webUrl?: string;
  sourceAttachmentId?: string;
}

// ---------------------------------------------------------------------------
// Tool definitions

export function withAccount(
  ctx: UtilityToolContext,
  schema: Record<string, z.ZodTypeAny>
): Record<string, z.ZodTypeAny> {
  if (ctx.multiAccount) {
    schema['account'] = z
      .string()
      .optional()
      .describe(
        'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
      );
  }
  return schema;
}

export function str(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined;
}

export const FILE_OPS_TOOLS: UtilityTool[] = [
  {
    name: 'save-attachment-to-drive',
    method: 'PUT',
    path: 'tool:save-attachment-to-drive',
    searchKeywords:
      'save attachment to onedrive sharepoint matter folder file email attachment upload copy attachment to drive',
    description:
      'Save a mail attachment into a OneDrive or SharePoint folder, server side (the bytes never pass through the conversation). Give the folder as folderItemId or as folderPath relative to the drive root. conflictBehavior defaults to "rename" (keeps both); "fail" refuses if the name exists; "replace" overwrites and is used only when passed explicitly. fileName defaults to the attachment name with SharePoint-illegal characters removed. Returns { itemId, name, size, webUrl }. Limit 150 MB.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        messageId: z.string().describe('Id of the message that carries the attachment.'),
        attachmentId: z.string().describe('Attachment id (list-mail-attachments returns it).'),
        driveId: z.string().describe('Destination drive id.'),
        folderItemId: z
          .string()
          .optional()
          .describe('Destination folder item id. Give this or folderPath, not both.'),
        folderPath: z
          .string()
          .optional()
          .describe(
            'Destination folder path relative to the drive root, e.g. "GC Associates/Matters/Client/Matter". Give this or folderItemId, not both.'
          ),
        fileName: z
          .string()
          .optional()
          .describe('File name to save as. Defaults to the attachment name, made SharePoint-safe.'),
        conflictBehavior: z
          .enum(CONFLICT_BEHAVIORS)
          .optional()
          .describe(
            'What to do when a file with that name already exists: rename (default), fail, or replace (overwrite; only when explicitly intended).'
          ),
      }),
    execute: async (params, ctx) => {
      const tool = 'save-attachment-to-drive';
      const messageId = str(params.messageId);
      const attachmentId = str(params.attachmentId);
      const driveId = str(params.driveId);
      const folderItemId = str(params.folderItemId);
      const folderPath = str(params.folderPath);
      if (!messageId || !attachmentId || !driveId) {
        return fail('messageId, attachmentId, and driveId are required.');
      }
      if (!folderItemId && !folderPath) return fail('Give folderItemId or folderPath.');
      if (folderItemId && folderPath) {
        return fail('Give folderItemId or folderPath, not both.');
      }
      const conflictRaw = params.conflictBehavior ?? 'rename';
      if (!CONFLICT_BEHAVIORS.includes(conflictRaw as ConflictBehavior)) {
        return fail('conflictBehavior must be "rename", "fail", or "replace".');
      }
      const conflictBehavior = conflictRaw as ConflictBehavior;
      const destination = folderItemId
        ? `drive:${driveId}/item:${folderItemId}`
        : `drive:${driveId}/path:${folderPath}`;
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        const file = await readMailAttachment(ctx.graphClient, token, messageId, attachmentId);
        const fileName = sharePointSafeName(str(params.fileName) ?? file.name);
        if (!fileName) return fail('The file name is empty after removing illegal characters.');
        let written: DriveWriteResult;
        try {
          written = await writeToDrive(
            ctx.graphClient,
            token,
            driveId,
            { itemId: folderItemId, path: folderPath },
            fileName,
            file.bytes,
            conflictBehavior
          );
        } catch (error) {
          if (statusOf(error) === 409) {
            throw new FileOpsError(
              `A file named "${fileName}" already exists in that folder and conflictBehavior is "${conflictBehavior}". Nothing was written.`
            );
          }
          throw new FileOpsError(
            describe(error, `Destination folder ${folderItemId ?? folderPath}`)
          );
        }
        logOp(tool, file.label, `${destination}/${written.name ?? fileName}`, file.bytes.length);
        return ok({
          itemId: written.id,
          name: written.name ?? fileName,
          size: written.size ?? file.bytes.length,
          webUrl: written.webUrl,
          conflictBehavior,
        });
      } catch (error) {
        logOp(
          tool,
          `message:${messageId}/attachment:${attachmentId}`,
          `${destination} FAILED`,
          null
        );
        return fail(describe(error, 'The item'));
      }
    },
  },
  {
    name: 'attach-drive-file-to-draft',
    method: 'POST',
    path: 'tool:attach-drive-file-to-draft',
    searchKeywords:
      'attach onedrive file to draft attach sharepoint document to email draft add attachment from drive',
    description:
      'Attach a OneDrive or SharePoint file to a draft email, server side (the bytes never pass through the conversation). Files under 3 MB are attached directly; 3 MB to 150 MB go through an attachment upload session. The target must be a draft. Returns { attachmentId, name, size }.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        draftMessageId: z.string().describe('Id of the draft message to attach to.'),
        driveId: z.string().describe('Drive id of the file.'),
        itemId: z.string().describe('Drive item id of the file.'),
        attachmentName: z
          .string()
          .optional()
          .describe('Name the attachment shows on the draft. Defaults to the file name.'),
      }),
    execute: async (params, ctx) => {
      const tool = 'attach-drive-file-to-draft';
      const draftId = str(params.draftMessageId);
      const driveId = str(params.driveId);
      const itemId = str(params.itemId);
      if (!draftId || !driveId || !itemId) {
        return fail('draftMessageId, driveId, and itemId are required.');
      }
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        const draft = await requireDraft(ctx.graphClient, token, draftId);
        const file = await readDriveFile(ctx.graphClient, token, driveId, itemId);
        const result = await attachToDraft(
          ctx.graphClient,
          token,
          draftId,
          file,
          str(params.attachmentName)
        ).catch((error) => {
          throw new FileOpsError(describe(error, `Draft message ${draftId}`));
        });
        logOp(tool, file.label, `draft:${draftId}/attachment:${result.attachmentId}`, result.size);
        return ok({
          attachmentId: result.attachmentId,
          name: result.name,
          size: result.size,
          uploadMethod: result.method,
          draftWebLink: draft.webLink,
        });
      } catch (error) {
        logOp(tool, `drive:${driveId}/item:${itemId}`, `draft:${draftId} FAILED`, null);
        return fail(describe(error, 'The item'));
      }
    },
  },
  {
    name: 'attach-mail-attachment-to-draft',
    method: 'POST',
    path: 'tool:attach-mail-attachment-to-draft',
    searchKeywords:
      'copy attachment from one email to a draft forward attachment to draft attach received attachment',
    description:
      'Copy an attachment from one message onto a draft email, server side (the bytes never pass through the conversation). Under 3 MB it is attached directly; 3 MB to 150 MB through an attachment upload session. The target must be a draft. Returns { attachmentId, name, size }.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        sourceMessageId: z.string().describe('Id of the message that carries the attachment.'),
        attachmentId: z.string().describe('Attachment id on the source message.'),
        draftMessageId: z.string().describe('Id of the draft message to attach to.'),
        attachmentName: z
          .string()
          .optional()
          .describe('Name the attachment shows on the draft. Defaults to the source name.'),
      }),
    execute: async (params, ctx) => {
      const tool = 'attach-mail-attachment-to-draft';
      const sourceId = str(params.sourceMessageId);
      const attachmentId = str(params.attachmentId);
      const draftId = str(params.draftMessageId);
      if (!sourceId || !attachmentId || !draftId) {
        return fail('sourceMessageId, attachmentId, and draftMessageId are required.');
      }
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        const draft = await requireDraft(ctx.graphClient, token, draftId);
        const file = await readMailAttachment(ctx.graphClient, token, sourceId, attachmentId);
        const result = await attachToDraft(
          ctx.graphClient,
          token,
          draftId,
          file,
          str(params.attachmentName)
        ).catch((error) => {
          throw new FileOpsError(describe(error, `Draft message ${draftId}`));
        });
        logOp(tool, file.label, `draft:${draftId}/attachment:${result.attachmentId}`, result.size);
        return ok({
          attachmentId: result.attachmentId,
          name: result.name,
          size: result.size,
          uploadMethod: result.method,
          draftWebLink: draft.webLink,
        });
      } catch (error) {
        logOp(
          tool,
          `message:${sourceId}/attachment:${attachmentId}`,
          `draft:${draftId} FAILED`,
          null
        );
        return fail(describe(error, 'The item'));
      }
    },
  },
  {
    name: 'find-matter-folder',
    method: 'GET',
    path: 'tool:find-matter-folder',
    searchKeywords: 'find matter folder client folder sharepoint matters locate case folder',
    description:
      'Find a client or matter folder under the firm Matters root by case-insensitive substring match on folder names. Searches two levels (client, then matter) and a third level under grouping folders (MATTERS_GROUPING_FOLDERS, default Other, Customers, Litigation, Insurance Claims) and under folders that matched the query, so Matters/Other/Bodec/NBU Public Information Request is found. Returns [{ group, client, matter, itemId, driveId, path, mattersPath, depth, webUrl, exact, children? }], exact matches first. path is from the drive root (usable as folderPath); mattersPath is from the Matters root. An exact match also lists its child folders [{ name, itemId, path }].',
    readOnlyHint: true,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        query: z.string().describe('Part of a client or matter folder name.'),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(50)
          .optional()
          .describe('Most results to return (default 10, max 50).'),
      }),
    execute: async (params, ctx) => {
      const tool = 'find-matter-folder';
      const query = str(params.query)?.toLowerCase();
      if (!query) return fail('query is required.');
      const maxRaw = params.maxResults ?? 10;
      if (typeof maxRaw !== 'number' || !Number.isInteger(maxRaw) || maxRaw < 1 || maxRaw > 50) {
        return fail('maxResults must be an integer from 1 to 50.');
      }
      const { driveId, rootPath } = mattersRoot();
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        const cacheKey = `${driveId}|${rootPath}`;
        let root = rootCache.get(cacheKey);
        if (!root) {
          let item: FolderChild;
          try {
            item = await graphJson<FolderChild>(
              ctx.graphClient,
              token,
              `/drives/${enc(driveId)}/root:/${encodePath(rootPath)}?$select=id,name,webUrl,folder`
            );
          } catch (error) {
            throw new FileOpsError(
              describe(error, `Matters root "${rootPath}" in drive ${driveId}`)
            );
          }
          if (!item.id) throw new FileOpsError(`Matters root "${rootPath}" returned no id.`);
          root = { id: item.id, webUrl: item.webUrl };
          rootCache.set(cacheKey, root);
        }
        const grouping = mattersGroupingFolders();
        const hit = (f: FolderChild) => (f.name ?? '').toLowerCase().includes(query);
        const isExact = (f: FolderChild) => (f.name ?? '').toLowerCase() === query;
        const list = (f: FolderChild) =>
          listChildFolders(ctx.graphClient, token, driveId, f.id as string);

        // Levels 1 and 2: every client folder and its children (the original two-level scan).
        const level1 = await listChildFolders(ctx.graphClient, token, driveId, root.id);
        const level2 = await mapLimit(level1, 8, list);
        const nodes: TreeNode[] = [];
        const childrenOf = new Map<string, FolderChild[]>();
        const pairs: Array<{ l1: FolderChild; l2: FolderChild }> = [];
        level1.forEach((l1, i) => {
          nodes.push({ segments: [l1] });
          childrenOf.set(l1.id as string, level2[i]);
          for (const l2 of level2[i]) {
            nodes.push({ segments: [l1, l2] });
            pairs.push({ l1, l2 });
          }
        });

        // Level 3: only under grouping folders and under folders that matched the query.
        const descend = pairs.filter(
          ({ l1, l2 }) => grouping.has((l1.name ?? '').toLowerCase()) || hit(l1) || hit(l2)
        );
        const level3 = await mapLimit(descend, 8, ({ l2 }) => list(l2));
        descend.forEach(({ l1, l2 }, i) => {
          childrenOf.set(l2.id as string, level3[i]);
          for (const l3 of level3[i]) nodes.push({ segments: [l1, l2, l3] });
        });

        const matched = nodes.filter((n) => hit(n.segments[n.segments.length - 1]));
        const matches: MatterMatch[] = matched.map((n) => ({
          ...describeNode(n, grouping, driveId, rootPath),
          exact: isExact(n.segments[n.segments.length - 1]),
        }));
        matches.sort((a, b) =>
          a.exact !== b.exact ? (a.exact ? -1 : 1) : a.path.localeCompare(b.path)
        );
        // Exact matches carry their child folders; a level-3 folder's children cost one call.
        const exactHits = matches.slice(0, maxRaw).filter((m) => m.exact);
        await mapLimit(exactHits, 8, async (m) => {
          let kids = childrenOf.get(m.itemId);
          if (!kids) kids = await list({ id: m.itemId });
          m.children = kids
            .map((k) => ({
              name: k.name ?? '',
              itemId: k.id as string,
              path: `${m.path}/${k.name ?? ''}`,
            }))
            .sort((a, b) => a.name.localeCompare(b.name));
        });
        const results = matches.slice(0, maxRaw);
        logOp(tool, `drive:${driveId}/path:${rootPath}`, `query:${query}`, matches.length);
        return ok(results);
      } catch (error) {
        logOp(tool, `drive:${driveId}/path:${rootPath}`, `query:${query} FAILED`, null);
        return fail(describe(error, 'The Matters folder'));
      }
    },
  },
  {
    name: 'file-email-to-matter',
    method: 'PUT',
    path: 'tool:file-email-to-matter',
    searchKeywords:
      'file email to matter folder save email as eml save message and attachments to sharepoint matter',
    description:
      'File an email into a matter folder, server side. Saves the full message (MIME) as <folder>/Emails/<yyyy-mm-dd> <subject>.eml (Denver date of receipt, Re/Fw prefixes and SharePoint-illegal characters removed, name capped at 120 characters), creating the Emails folder if missing, then saves each file attachment into the folder itself. Never overwrites: a name that already exists is saved as a renamed copy. Inline images (signature logos) and cloud-link attachments are skipped unless includeInlineAttachments is true; skipped items are listed. Returns { items: [{ kind, itemId, name, size, webUrl }], skipped, failed, emailsFolder }.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        messageId: z.string().describe('Id of the message to file.'),
        driveId: z.string().describe('Drive id of the matter folder.'),
        folderItemId: z
          .string()
          .describe('Item id of the matter folder (find-matter-folder returns it).'),
        saveEml: z
          .boolean()
          .optional()
          .describe('Save the message as an .eml in the Emails sub-folder (default true).'),
        saveAttachments: z
          .boolean()
          .optional()
          .describe('Save each file attachment into the matter folder (default true).'),
        includeInlineAttachments: z
          .boolean()
          .optional()
          .describe('Also save inline attachments such as signature images (default false).'),
      }),
    execute: async (params, ctx) => {
      const tool = 'file-email-to-matter';
      const messageId = str(params.messageId);
      const driveId = str(params.driveId);
      const folderItemId = str(params.folderItemId);
      if (!messageId || !driveId || !folderItemId) {
        return fail('messageId, driveId, and folderItemId are required.');
      }
      const saveEml = params.saveEml ?? true;
      const saveAttachments = params.saveAttachments ?? true;
      const includeInline = params.includeInlineAttachments ?? false;
      if (
        typeof saveEml !== 'boolean' ||
        typeof saveAttachments !== 'boolean' ||
        typeof includeInline !== 'boolean'
      ) {
        return fail(
          'saveEml, saveAttachments, and includeInlineAttachments must be true or false.'
        );
      }
      if (!saveEml && !saveAttachments) {
        return fail('Nothing to do: saveEml and saveAttachments are both false.');
      }
      const source = `message:${messageId}`;
      const destination = `drive:${driveId}/item:${folderItemId}`;
      const items: FiledItem[] = [];
      const skipped: Array<{ name: string; reason: string }> = [];
      const failed: Array<{ name: string; error: string }> = [];
      let emailsFolder: { itemId: string; webUrl?: string; created: boolean } | null = null;
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        let message: { subject?: string; receivedDateTime?: string; sentDateTime?: string };
        try {
          message = await graphJson(
            ctx.graphClient,
            token,
            `/me/messages/${enc(messageId)}?$select=id,subject,receivedDateTime,sentDateTime`
          );
        } catch (error) {
          throw new FileOpsError(describe(error, `Message ${messageId}`));
        }

        if (saveEml) {
          let folder: { id: string; webUrl?: string; created: boolean };
          try {
            folder = await ensureChildFolder(
              ctx.graphClient,
              token,
              driveId,
              folderItemId,
              EMAILS_FOLDER_NAME
            );
          } catch (error) {
            throw new FileOpsError(describe(error, `Matter folder ${folderItemId}`));
          }
          emailsFolder = { itemId: folder.id, webUrl: folder.webUrl, created: folder.created };
          let mime: Buffer;
          try {
            mime = await graphBytes(
              ctx.graphClient,
              token,
              `/me/messages/${enc(messageId)}/$value`
            );
          } catch (error) {
            throw new FileOpsError(describe(error, `Message ${messageId} content`));
          }
          const name = emlFileName(
            message.subject,
            firmDate(message.receivedDateTime ?? message.sentDateTime)
          );
          try {
            const written = await writeToDrive(
              ctx.graphClient,
              token,
              driveId,
              { itemId: folder.id },
              name,
              mime,
              'rename'
            );
            items.push({
              kind: 'eml',
              itemId: written.id,
              name: written.name ?? name,
              size: written.size ?? mime.length,
              webUrl: written.webUrl,
            });
          } catch (error) {
            throw new FileOpsError(describe(error, `Emails folder ${folder.id}`));
          }
        }

        if (saveAttachments) {
          let rows: MessageAttachmentRow[];
          try {
            rows = await listMessageAttachments(ctx.graphClient, token, messageId);
          } catch (error) {
            throw new FileOpsError(describe(error, `Attachments of message ${messageId}`));
          }
          for (const row of rows) {
            const label = row.name ?? row.id ?? 'attachment';
            const kind = (row['@odata.type'] ?? '').toLowerCase();
            if (kind.endsWith('referenceattachment')) {
              skipped.push({
                name: label,
                reason: 'cloud link (reference attachment), not a file',
              });
              continue;
            }
            if (row.isInline && !includeInline) {
              skipped.push({ name: label, reason: 'inline (embedded in the message body)' });
              continue;
            }
            try {
              const file = await readMailAttachment(
                ctx.graphClient,
                token,
                messageId,
                row.id as string
              );
              const fileName = sharePointSafeName(file.name);
              if (!fileName) throw new FileOpsError('The file name is empty after cleaning.');
              const written = await writeToDrive(
                ctx.graphClient,
                token,
                driveId,
                { itemId: folderItemId },
                fileName,
                file.bytes,
                'rename'
              );
              items.push({
                kind: 'attachment',
                itemId: written.id,
                name: written.name ?? fileName,
                size: written.size ?? file.bytes.length,
                webUrl: written.webUrl,
                sourceAttachmentId: row.id,
              });
            } catch (error) {
              failed.push({ name: label, error: describe(error, `Attachment "${label}"`) });
            }
          }
        }

        const total = items.reduce((n, i) => n + i.size, 0);
        logOp(
          tool,
          source,
          `${destination} written:${items.length} skipped:${skipped.length} failed:${failed.length}`,
          total
        );
        const payload = { items, skipped, failed, emailsFolder };
        if (items.length === 0 && failed.length > 0) {
          return {
            content: [
              { type: 'text', text: JSON.stringify({ error: 'Nothing was written.', ...payload }) },
            ],
            isError: true,
          };
        }
        return ok(payload);
      } catch (error) {
        logOp(tool, source, `${destination} FAILED written:${items.length}`, null);
        const message = describe(error, 'The item');
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: message, items, skipped, failed, emailsFolder }),
            },
          ],
          isError: true,
        };
      }
    },
  },
];
