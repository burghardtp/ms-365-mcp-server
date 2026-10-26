/**
 * Server-side send gate (firm design, section E).
 *
 * Every tool that sends mail or shares a file needs a gateRecordId from
 * send_gate_record on Paul Context. Before the Graph call the server POSTs
 * {id, tool, connector, target} to GATE_CHECK_URL with Bearer SEND_GATE_TOKEN.
 * Anything other than an ok JSON answer refuses the send, and no Graph call is
 * made. With GATE_CHECK_URL unset every send is refused (fail closed) unless
 * GATE_DISABLED=true, which is for local tests only.
 *
 * send-draft-message also reads the draft and its attachment list first: a body
 * that says something is attached, with no file attachment on the draft, is
 * refused before the gate is asked. The recipients, subject, and attachment
 * names and sizes go into the gate target and into the tool's success response.
 */
import type GraphClient from './graph-client.js';
import logger from './logger.js';

/**
 * The endpoints.json fields that turn the gate on. sendGate marks a send or share
 * tool; draftAttachmentCheck adds the read of the draft and its attachments.
 */
export interface SendGateConfig {
  sendGate?: boolean;
  draftAttachmentCheck?: boolean;
}

export const GATE_RECORD_PARAM = 'gateRecordId';

export const GATE_RECORD_PARAM_DESCRIPTION =
  'Send gate record id from send_gate_record on Paul Context. Required; the send is refused without a valid, unexpired record.';

export const ATTACHMENT_MENTION_PATTERN = /\battach(ed|ment|ments|ing)?\b|\benclos(ed|ure|ures)\b/i;

export const ATTACHMENT_MISSING_MESSAGE =
  'The draft says something is attached but carries no attachment.';

export const GATE_NOT_CONFIGURED_MESSAGE = 'send gate is not configured on this server';

/** A text-only tool result, assignable to the MCP CallToolResult shapes in use. */
export interface GateRefusal {
  content: Array<{ type: 'text'; text: string }>;
  isError: true;
  [key: string]: unknown;
}

const TARGET_MAX = 300;
const GATE_TIMEOUT_MS = 15_000;

export function isSendGated(config: SendGateConfig | undefined): boolean {
  return config?.sendGate === true;
}

let configuredPublicUrl: string | undefined;

/** Called once by the server with its public URL, the fallback connector name. */
export function configureSendGate(options: { publicUrl?: string | null }): void {
  configuredPublicUrl = options.publicUrl ?? undefined;
}

export function connectorName(): string {
  const explicit = process.env.GATE_CONNECTOR_NAME?.trim();
  if (explicit) return explicit;
  const raw =
    configuredPublicUrl || process.env.MS365_MCP_PUBLIC_URL || process.env.MS365_MCP_BASE_URL;
  if (raw) {
    try {
      return new URL(raw).host;
    } catch {
      // fall through
    }
  }
  return 'ms-365-mcp-server';
}

export interface AttachmentFact {
  name: string;
  size: number;
  isInline: boolean;
  contentType?: string;
}

export interface SendFacts {
  to: string[];
  cc: string[];
  bcc: string[];
  subject?: string;
  attachments?: AttachmentFact[];
}

export interface GateApproval {
  id: string;
  matter?: string;
  doc?: string;
  ts?: number;
}

export type GateOutcome =
  | { ok: true; approval: GateApproval | null; facts: SendFacts; target: string }
  | { ok: false; result: GateRefusal; facts: SendFacts };

function refusal(tool: string, reason: string, extra: Record<string, unknown> = {}): GateRefusal {
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({ error: 'send_gate_refused', tool, reason, ...extra }),
      },
    ],
    isError: true,
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function readAddress(entry: unknown): string | undefined {
  if (typeof entry === 'string') return entry;
  if (!isRecord(entry)) return undefined;
  const email = entry.emailAddress ?? entry.EmailAddress;
  if (isRecord(email) && typeof email.address === 'string') return email.address;
  if (typeof entry.email === 'string') return entry.email;
  if (typeof entry.address === 'string') return entry.address;
  return undefined;
}

function addresses(list: unknown): string[] {
  if (!Array.isArray(list)) return [];
  return list.map(readAddress).filter((a): a is string => Boolean(a));
}

/**
 * Recipients and subject from a send tool's arguments. The body may sit under
 * `body`, under `body.message` (send-mail, reply), or be flattened into the
 * top-level params; every location is read.
 */
export function factsFromParams(params: Record<string, unknown>): SendFacts {
  const facts: SendFacts = { to: [], cc: [], bcc: [] };
  const containers: Record<string, unknown>[] = [params];
  const queue: unknown[] = [params.body, params.message, params.Message];
  while (queue.length > 0) {
    const next = queue.shift();
    if (!isRecord(next)) continue;
    containers.push(next);
    queue.push(next.message, next.Message);
  }
  for (const c of containers) {
    facts.to.push(...addresses(c.toRecipients), ...addresses(c.recipients));
    facts.cc.push(...addresses(c.ccRecipients));
    facts.bcc.push(...addresses(c.bccRecipients));
    if (!facts.subject && typeof c.subject === 'string' && c.subject) facts.subject = c.subject;
  }
  facts.to = [...new Set(facts.to)];
  facts.cc = [...new Set(facts.cc)];
  facts.bcc = [...new Set(facts.bcc)];
  return facts;
}

function param(params: Record<string, unknown>, ...names: string[]): string | undefined {
  for (const n of names) {
    const v = params[n];
    if (typeof v === 'string' && v) return v;
  }
  return undefined;
}

export function messageIdParam(params: Record<string, unknown>): string | undefined {
  return param(params, 'messageId', 'message-id');
}

export function buildTarget(
  tool: string,
  params: Record<string, unknown>,
  facts: SendFacts
): string {
  const parts: string[] = [];
  // Shared mailbox sends carry the mailbox as a path param.
  const mailbox = param(params, 'userId', 'user-id');
  if (mailbox) parts.push(`mailbox: ${mailbox}`);
  if (facts.to.length) parts.push(`to: ${facts.to.join(', ')}`);
  if (facts.cc.length) parts.push(`cc: ${facts.cc.join(', ')}`);
  if (facts.bcc.length) parts.push(`bcc: ${facts.bcc.join(', ')}`);
  if (facts.subject) parts.push(`subject: ${facts.subject}`);
  if (facts.attachments) {
    const files = facts.attachments.filter((a) => !a.isInline);
    parts.push(
      files.length
        ? `attachments: ${files.map((a) => `${a.name} (${a.size} bytes)`).join(', ')}`
        : 'attachments: none'
    );
  }
  // File shares carry the drive item as path params; message sends the message id.
  const drive = param(params, 'driveId', 'drive-id');
  const item = param(params, 'driveItemId', 'driveItem-id');
  if (drive || item) parts.push(`drive ${drive ?? '?'} item ${item ?? '?'}`);
  const id = messageIdParam(params);
  if (id) parts.push(`message ${id}`);
  const target = parts.join('; ') || tool;
  return target.length > TARGET_MAX ? target.slice(0, TARGET_MAX) : target;
}

const ENTITIES: Record<string, string> = {
  nbsp: ' ',
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Plain text of an HTML (or text) message body, for the attachment-mention test. */
export function bodyText(content: string, contentType?: string): string {
  if (contentType && contentType.toLowerCase() === 'text') return content;
  return content
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(style|script|head)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
      const lower = e.toLowerCase();
      if (lower.startsWith('#x')) return String.fromCodePoint(parseInt(lower.slice(2), 16));
      if (lower.startsWith('#')) return String.fromCodePoint(parseInt(lower.slice(1), 10));
      return ENTITIES[lower] ?? m;
    })
    .replace(/\s+/g, ' ')
    .trim();
}

interface DraftMessage {
  subject?: string;
  toRecipients?: unknown;
  ccRecipients?: unknown;
  bccRecipients?: unknown;
  body?: { content?: string; contentType?: string };
  hasAttachments?: boolean;
}

interface AttachmentList {
  value?: Array<{
    id?: string;
    name?: string;
    size?: number;
    isInline?: boolean;
    contentType?: string;
  }>;
}

function graphErrorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Read the draft and its attachments for send-draft-message. */
export async function readDraftFacts(
  graph: GraphClient,
  messageId: string,
  accessToken: string | undefined
): Promise<{ facts: SendFacts; text: string }> {
  const id = encodeURIComponent(messageId).replace(/%3D/g, '=');
  const opts = { method: 'GET', ...(accessToken ? { accessToken } : {}) };
  const draft = (await graph.makeRequest(
    `/me/messages/${id}?$select=subject,toRecipients,ccRecipients,bccRecipients,body,hasAttachments`,
    opts
  )) as DraftMessage;
  const list = (await graph.makeRequest(
    `/me/messages/${id}/attachments?$select=id,name,size,isInline,contentType`,
    opts
  )) as AttachmentList;
  const attachments: AttachmentFact[] = (list?.value ?? []).map((a) => ({
    name: a.name ?? '(unnamed)',
    size: typeof a.size === 'number' ? a.size : 0,
    isInline: a.isInline === true,
    ...(a.contentType ? { contentType: a.contentType } : {}),
  }));
  const facts: SendFacts = {
    to: addresses(draft?.toRecipients),
    cc: addresses(draft?.ccRecipients),
    bcc: addresses(draft?.bccRecipients),
    ...(draft?.subject ? { subject: draft.subject } : {}),
    attachments,
  };
  return { facts, text: bodyText(draft?.body?.content ?? '', draft?.body?.contentType) };
}

async function askGate(
  url: string,
  token: string,
  payload: { id: string; tool: string; connector: string; target: string }
): Promise<{ ok: true; approval: GateApproval } | { ok: false; reason: string }> {
  let response: Response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(GATE_TIMEOUT_MS),
    });
  } catch (error) {
    return { ok: false, reason: `Send gate check failed: ${graphErrorText(error)}` };
  }
  let data: unknown;
  try {
    data = await response.json();
  } catch {
    data = undefined;
  }
  const reason = isRecord(data) && typeof data.reason === 'string' ? data.reason : undefined;
  if (!response.ok) {
    return {
      ok: false,
      reason: reason ?? `Send gate check answered HTTP ${response.status}.`,
    };
  }
  if (!isRecord(data) || data.ok !== true) {
    return { ok: false, reason: reason ?? 'Send gate check did not answer ok.' };
  }
  return {
    ok: true,
    approval: {
      id: typeof data.id === 'string' ? data.id : payload.id,
      ...(typeof data.matter === 'string' ? { matter: data.matter } : {}),
      ...(typeof data.doc === 'string' ? { doc: data.doc } : {}),
      ...(typeof data.ts === 'number' ? { ts: data.ts } : {}),
    },
  };
}

/**
 * Run the gate for one send. Returns ok with the facts to echo, or a ready
 * refusal. Never calls Graph's send endpoint itself.
 */
export async function runSendGate(
  tool: string,
  config: SendGateConfig | undefined,
  params: Record<string, unknown>,
  graph: GraphClient,
  accessToken: string | undefined
): Promise<GateOutcome> {
  const rawId = params[GATE_RECORD_PARAM];
  const gateId = typeof rawId === 'string' ? rawId.trim() : '';
  let facts = factsFromParams(params);

  // Configuration and the record id are checked first, so a send refused for
  // either reason touches nothing in Graph.
  const url = process.env.GATE_CHECK_URL?.trim();
  const disabled = !url && process.env.GATE_DISABLED === 'true';
  const token = process.env.SEND_GATE_TOKEN?.trim();
  if (!disabled) {
    if (!url) {
      return { ok: false, facts, result: refusal(tool, GATE_NOT_CONFIGURED_MESSAGE) };
    }
    if (!token) {
      return {
        ok: false,
        facts,
        result: refusal(tool, `${GATE_NOT_CONFIGURED_MESSAGE} (SEND_GATE_TOKEN is not set)`),
      };
    }
    if (!gateId) {
      return {
        ok: false,
        facts,
        result: refusal(
          tool,
          'No gateRecordId was given. Call send_gate_record on Paul Context and pass its id as gateRecordId.'
        ),
      };
    }
  }

  if (config?.draftAttachmentCheck === true) {
    const messageId = messageIdParam(params);
    if (!messageId) {
      return { ok: false, facts, result: refusal(tool, 'messageId is required.') };
    }
    let draft: { facts: SendFacts; text: string };
    try {
      draft = await readDraftFacts(graph, messageId, accessToken);
    } catch (error) {
      return {
        ok: false,
        facts,
        result: refusal(tool, `Could not read the draft before sending: ${graphErrorText(error)}`),
      };
    }
    facts = draft.facts;
    const files = (facts.attachments ?? []).filter((a) => !a.isInline);
    if (ATTACHMENT_MENTION_PATTERN.test(draft.text) && files.length === 0) {
      return {
        ok: false,
        facts,
        result: refusal(tool, ATTACHMENT_MISSING_MESSAGE, {
          to: facts.to,
          cc: facts.cc,
          bcc: facts.bcc,
          subject: facts.subject ?? '',
          attachments: [],
        }),
      };
    }
  }

  const target = buildTarget(tool, params, facts);
  if (disabled || !url || !token) {
    return { ok: true, approval: null, facts, target };
  }

  const answer = await askGate(url, token, {
    id: gateId,
    tool,
    connector: connectorName(),
    target,
  });
  if (!answer.ok) {
    return { ok: false, facts, result: refusal(tool, answer.reason, { gateRecordId: gateId }) };
  }
  return { ok: true, approval: answer.approval, facts, target };
}

/** The success echo appended to the tool result. */
export function sendSummary(
  tool: string,
  approval: GateApproval | null,
  facts: SendFacts
): Record<string, unknown> {
  return {
    sendGate: approval
      ? { id: approval.id, matter: approval.matter, doc: approval.doc }
      : { disabled: true },
    tool,
    to: facts.to,
    cc: facts.cc,
    bcc: facts.bcc,
    ...(facts.subject !== undefined ? { subject: facts.subject } : {}),
    ...(facts.attachments
      ? {
          attachments: facts.attachments
            .filter((a) => !a.isInline)
            .map((a) => ({ name: a.name, size: a.size })),
        }
      : {}),
  };
}

/** One info line per send attempt. Never carries a token. */
export function logSend(tool: string, gateId: string, facts: SendFacts, result: string): void {
  const recipients = facts.to.length + facts.cc.length + facts.bcc.length;
  const attachments = facts.attachments
    ? facts.attachments.filter((a) => !a.isInline).length
    : 'n/a';
  logger.info(
    `send-gate ts=${new Date().toISOString()} tool=${tool} gate=${gateId || '(none)'} recipients=${recipients} attachments=${attachments} result=${result}`
  );
}
