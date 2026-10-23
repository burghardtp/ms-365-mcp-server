import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const { logLines } = vi.hoisted(() => ({ logLines: [] as string[] }));
vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn((m: string) => logLines.push(m)),
    error: vi.fn(),
    warn: vi.fn(),
    debug: vi.fn(),
  },
}));

import GraphClient from '../src/graph-client.js';
import type AuthManager from '../src/auth.js';
import type { AppSecrets } from '../src/secrets.js';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import {
  __resetMattersRootCacheForTests,
  DEFAULT_MATTERS_DRIVE_ID,
  FILE_OPS_MAX_BYTES,
  sharePointSafeName,
  UPLOAD_CHUNK_BYTES,
} from '../src/file-ops.js';
import { __resetSharedBreakerForTests } from '../src/lib/graph-resilience.js';
import { TOOL_CATEGORIES } from '../src/tool-categories.js';

const G = 'https://graph.microsoft.com/v1.0';
const UPLOAD_URL = 'https://upload.example.test/session/abc';

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: Buffer | string | undefined;
}

type Handler = (call: Call) => Response | undefined;

let calls: Call[] = [];
let handlers: Handler[] = [];

function on(method: string, match: string | RegExp, respond: (call: Call) => Response): void {
  handlers.push((call) => {
    if (call.method !== method) return undefined;
    const hit = typeof match === 'string' ? call.url === match : match.test(call.url);
    return hit ? respond(call) : undefined;
  });
}

function json(data: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json', ...headers },
  });
}

function bytes(n: number, seed = 7): Buffer {
  const b = Buffer.alloc(n);
  for (let i = 0; i < n; i++) b[i] = (i * seed + 3) & 0xff;
  return b;
}

function binary(buf: Buffer, contentType = 'application/pdf'): Response {
  return new Response(buf, {
    status: 200,
    headers: { 'content-type': contentType, 'content-length': String(buf.length) },
  });
}

function graphError(status: number, code: string): Response {
  return json({ error: { code, message: code } }, status);
}

function makeCtx(token: string | null = 'test-token') {
  const authManager = {
    isOAuthModeEnabled: () => false,
    getToken: async () => token,
    getTokenForAccount: async () => {
      if (!token) throw new Error('no cached account');
      return token;
    },
  } as unknown as AuthManager;
  const graphClient = new GraphClient(authManager, { cloudType: 'global' } as AppSecrets);
  return { graphClient, authManager, multiAccount: false, accountNames: [] };
}

function tool(name: string) {
  const t = UTILITY_TOOLS.find((u) => u.name === name);
  if (!t) throw new Error(`tool ${name} not registered`);
  return t;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parse(result: { content: Array<{ text?: string }> }): any {
  return JSON.parse(result.content[0].text as string);
}

beforeEach(() => {
  calls = [];
  handlers = [];
  logLines.length = 0;
  __resetSharedBreakerForTests();
  __resetMattersRootCacheForTests();
  vi.stubGlobal(
    'fetch',
    vi.fn(async (input: string | URL, init: Parameters<typeof fetch>[1] = {}) => {
      const raw = init.headers as Record<string, string> | undefined;
      const headers: Record<string, string> = {};
      for (const [k, v] of Object.entries(raw ?? {})) headers[k.toLowerCase()] = v;
      const b = init.body;
      const call: Call = {
        url: String(input),
        method: (init.method ?? 'GET').toUpperCase(),
        headers,
        body:
          b === undefined || b === null
            ? undefined
            : typeof b === 'string'
              ? b
              : Buffer.from(b as Uint8Array),
      };
      calls.push(call);
      for (const h of handlers) {
        const r = h(call);
        if (r) return r;
      }
      return graphError(404, `unmocked ${call.method} ${call.url}`);
    })
  );
});

afterEach(() => {
  vi.unstubAllGlobals();
  delete process.env.MATTERS_DRIVE_ID;
  delete process.env.MATTERS_ROOT_PATH;
});

const ATT_META = `${G}/me/messages/MSG1/attachments/ATT1?$select=id,name,contentType,size`;
const ATT_VALUE = `${G}/me/messages/MSG1/attachments/ATT1/$value`;

function mockAttachment(buf: Buffer, name = 'Exhibit A.pdf', declaredSize = buf.length) {
  on('GET', ATT_META, () =>
    json({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'ATT1',
      name,
      contentType: 'application/pdf',
      size: declaredSize,
    })
  );
  on('GET', ATT_VALUE, () => binary(buf));
}

function noBase64(result: { content: Array<{ text?: string }> }, buf: Buffer): void {
  const text = result.content.map((c) => c.text ?? '').join('');
  expect(text).not.toContain(buf.toString('base64').slice(0, 64));
  expect(text).not.toContain('contentBytes');
}

describe('registration', () => {
  it('registers the four file tools and puts them in the presets', () => {
    for (const name of [
      'save_attachment_to_drive',
      'attach_drive_file_to_draft',
      'attach_mail_attachment_to_draft',
      'find_matter_folder',
    ]) {
      expect(UTILITY_TOOLS.some((u) => u.name === name)).toBe(true);
    }
    expect(TOOL_CATEGORIES.mail.pattern.test('save_attachment_to_drive')).toBe(true);
    expect(TOOL_CATEGORIES.mail.pattern.test('attach_mail_attachment_to_draft')).toBe(true);
    expect(TOOL_CATEGORIES.files.pattern.test('attach_drive_file_to_draft')).toBe(true);
    expect(TOOL_CATEGORIES.files.pattern.test('find_matter_folder')).toBe(true);
    expect(tool('find_matter_folder').readOnlyHint).toBe(true);
    expect(tool('save_attachment_to_drive').readOnlyHint).toBe(false);
  });
});

describe('sharePointSafeName', () => {
  it('strips the characters SharePoint refuses', () => {
    expect(sharePointSafeName('Re: Exhibit #1 <draft> "50%" a|b? *.pdf')).toBe(
      'Re Exhibit 1 draft 50 ab .pdf'
    );
    expect(sharePointSafeName('a/b\\c.pdf')).toBe('abc.pdf');
    expect(sharePointSafeName('  ..name.pdf. ')).toBe('name.pdf');
  });
});

describe('save_attachment_to_drive', () => {
  const t = () => tool('save_attachment_to_drive');

  it('small file: single PUT with conflictBehavior=rename by default and a safe name', async () => {
    const buf = bytes(1000);
    mockAttachment(buf, 'Exhibit: A?.pdf');
    on('PUT', /\/content\?/, (call) => {
      expect(call.headers['content-type']).toBe('application/octet-stream');
      expect((call.body as Buffer).equals(buf)).toBe(true);
      return json({ id: 'NEW1', name: 'Exhibit A.pdf', size: 1000, webUrl: 'https://sp/x' }, 201);
    });
    const r = await t().execute(
      { messageId: 'MSG1', attachmentId: 'ATT1', driveId: 'D1', folderItemId: 'F1' },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    expect(parse(r)).toMatchObject({
      itemId: 'NEW1',
      name: 'Exhibit A.pdf',
      size: 1000,
      webUrl: 'https://sp/x',
      conflictBehavior: 'rename',
    });
    const put = calls.find((c) => c.method === 'PUT')!;
    expect(put.url).toBe(
      `${G}/drives/D1/items/F1:/Exhibit%20A.pdf:/content?@microsoft.graph.conflictBehavior=rename`
    );
    expect(put.headers.authorization).toBe('Bearer test-token');
    noBase64(r, buf);
    const log = logLines.find((l) => l.startsWith('[file-ops]'))!;
    const entry = JSON.parse(log.slice('[file-ops] '.length));
    expect(entry).toMatchObject({ tool: 'save_attachment_to_drive', size: 1000 });
    expect(entry.source).toContain('MSG1');
    expect(entry.destination).toContain('F1');
    expect(typeof entry.ts).toBe('string');
  });

  it('folderPath addresses the folder from the drive root', async () => {
    mockAttachment(bytes(10), 'a.pdf');
    on('PUT', /\/content\?/, () => json({ id: 'N', name: 'a.pdf', size: 10, webUrl: 'w' }, 201));
    const r = await t().execute(
      {
        messageId: 'MSG1',
        attachmentId: 'ATT1',
        driveId: 'D1',
        folderPath: '/GC Associates/Matters/BODEC/',
      },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    expect(calls.find((c) => c.method === 'PUT')!.url).toBe(
      `${G}/drives/D1/root:/GC%20Associates/Matters/BODEC/a.pdf:/content?@microsoft.graph.conflictBehavior=rename`
    );
  });

  it('large file: upload session carrying conflictBehavior, ranges without a bearer header', async () => {
    const size = 5 * 1024 * 1024;
    const buf = bytes(size);
    mockAttachment(buf, 'big.pdf');
    on('POST', `${G}/drives/D1/items/F1:/big.pdf:/createUploadSession`, (call) => {
      expect(JSON.parse(call.body as string)).toEqual({
        item: { '@microsoft.graph.conflictBehavior': 'fail', name: 'big.pdf' },
      });
      return json({ uploadUrl: UPLOAD_URL });
    });
    const received: Buffer[] = [];
    on('PUT', UPLOAD_URL, (call) => {
      received.push(call.body as Buffer);
      const total = received.reduce((n, b) => n + b.length, 0);
      return total < size
        ? json({ nextExpectedRanges: [`${total}-`] }, 202)
        : json({ id: 'BIG', name: 'big.pdf', size, webUrl: 'https://sp/big' }, 201);
    });
    const r = await t().execute(
      {
        messageId: 'MSG1',
        attachmentId: 'ATT1',
        driveId: 'D1',
        folderItemId: 'F1',
        conflictBehavior: 'fail',
      },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    expect(parse(r)).toMatchObject({ itemId: 'BIG', size, webUrl: 'https://sp/big' });
    const puts = calls.filter((c) => c.url === UPLOAD_URL);
    expect(puts).toHaveLength(Math.ceil(size / UPLOAD_CHUNK_BYTES));
    expect(puts[0].headers['content-range']).toBe(`bytes 0-${UPLOAD_CHUNK_BYTES - 1}/${size}`);
    expect(puts[1].headers['content-range']).toBe(
      `bytes ${UPLOAD_CHUNK_BYTES}-${size - 1}/${size}`
    );
    for (const p of puts) expect(p.headers.authorization).toBeUndefined();
    expect(Buffer.concat(received).equals(buf)).toBe(true);
    expect(calls.some((c) => c.method === 'PUT' && c.url.includes('/content'))).toBe(false);
  });

  it('conflictBehavior fail: a 409 is reported as a name clash and nothing else is tried', async () => {
    mockAttachment(bytes(10), 'a.pdf');
    on('PUT', /conflictBehavior=fail/, () => graphError(409, 'nameAlreadyExists'));
    const r = await tool('save_attachment_to_drive').execute(
      {
        messageId: 'MSG1',
        attachmentId: 'ATT1',
        driveId: 'D1',
        folderItemId: 'F1',
        conflictBehavior: 'fail',
      },
      makeCtx()
    );
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/already exists/);
    expect(calls.filter((c) => c.method === 'PUT')).toHaveLength(1);
  });

  it('replace is sent only when passed explicitly; unknown values are refused', async () => {
    mockAttachment(bytes(10), 'a.pdf');
    on('PUT', /\/content\?/, () => json({ id: 'N', name: 'a.pdf', size: 10, webUrl: 'w' }));
    await tool('save_attachment_to_drive').execute(
      {
        messageId: 'MSG1',
        attachmentId: 'ATT1',
        driveId: 'D1',
        folderItemId: 'F1',
        conflictBehavior: 'replace',
      },
      makeCtx()
    );
    expect(calls.find((c) => c.method === 'PUT')!.url).toMatch(/conflictBehavior=replace$/);
    calls = [];
    const bad = await tool('save_attachment_to_drive').execute(
      {
        messageId: 'MSG1',
        attachmentId: 'ATT1',
        driveId: 'D1',
        folderItemId: 'F1',
        conflictBehavior: 'overwrite',
      },
      makeCtx()
    );
    expect(bad.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it('refuses a missing folder argument, a missing token, and a missing attachment', async () => {
    const noFolder = await tool('save_attachment_to_drive').execute(
      { messageId: 'MSG1', attachmentId: 'ATT1', driveId: 'D1' },
      makeCtx()
    );
    expect(parse(noFolder).error).toMatch(/folderItemId or folderPath/);

    const noToken = await tool('save_attachment_to_drive').execute(
      { messageId: 'MSG1', attachmentId: 'ATT1', driveId: 'D1', folderItemId: 'F1' },
      makeCtx(null)
    );
    expect(noToken.isError).toBe(true);
    expect(parse(noToken).error).toMatch(/No Microsoft Graph access token/);
    expect(calls).toHaveLength(0);

    on('GET', ATT_META, () => graphError(404, 'ErrorItemNotFound'));
    const missing = await tool('save_attachment_to_drive').execute(
      { messageId: 'MSG1', attachmentId: 'ATT1', driveId: 'D1', folderItemId: 'F1' },
      makeCtx()
    );
    expect(parse(missing).error).toMatch(/Attachment ATT1 on message MSG1 was not found/);
  });

  it('refuses an attachment over 150 MB before downloading it', async () => {
    mockAttachment(bytes(10), 'huge.pdf', FILE_OPS_MAX_BYTES + 1);
    const r = await tool('save_attachment_to_drive').execute(
      { messageId: 'MSG1', attachmentId: 'ATT1', driveId: 'D1', folderItemId: 'F1' },
      makeCtx()
    );
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/over the 150 MB/);
    expect(calls.some((c) => c.url === ATT_VALUE)).toBe(false);
  });
});

const DRAFT = `${G}/me/messages/DRAFT1?$select=id,isDraft,subject,webLink`;
const ITEM_META = `${G}/drives/D1/items/I1?$select=id,name,size,webUrl,file,folder`;
const ITEM_CONTENT = `${G}/drives/D1/items/I1/content`;

function mockDraft(isDraft = true) {
  on('GET', DRAFT, () =>
    json({ id: 'DRAFT1', isDraft, subject: 'NBU response', webLink: 'https://outlook/d1' })
  );
}

function mockDriveFile(buf: Buffer, name = 'AG Letter.pdf', declaredSize = buf.length) {
  on('GET', ITEM_META, () =>
    json({ id: 'I1', name, size: declaredSize, file: { mimeType: 'application/pdf' } })
  );
  on('GET', ITEM_CONTENT, () => binary(buf));
}

function mockAttachmentUploadSession(size: number, received: Buffer[]) {
  on('POST', `${G}/me/messages/DRAFT1/attachments/createUploadSession`, () =>
    json({ uploadUrl: UPLOAD_URL })
  );
  on('PUT', UPLOAD_URL, (call) => {
    received.push(call.body as Buffer);
    const total = received.reduce((n, b) => n + b.length, 0);
    return total < size
      ? json({ nextExpectedRanges: [`${total}-`] }, 200)
      : new Response(null, {
          status: 201,
          headers: {
            Location:
              "https://outlook.office.com/api/v2.0/Users('u')/Messages('DRAFT1')/Attachments('NEWATT%3D')",
          },
        });
  });
}

describe('attach_drive_file_to_draft', () => {
  const t = () => tool('attach_drive_file_to_draft');

  it('small file: one POST fileAttachment built server side, no bytes returned', async () => {
    const buf = bytes(2000);
    mockDraft();
    mockDriveFile(buf);
    on('POST', `${G}/me/messages/DRAFT1/attachments`, (call) => {
      const body = JSON.parse(call.body as string);
      expect(body['@odata.type']).toBe('#microsoft.graph.fileAttachment');
      expect(body.name).toBe('AG Letter.pdf');
      expect(Buffer.from(body.contentBytes, 'base64').equals(buf)).toBe(true);
      return json(
        { id: 'A1', name: 'AG Letter.pdf', size: 2000, contentBytes: body.contentBytes },
        201
      );
    });
    const r = await t().execute(
      { draftMessageId: 'DRAFT1', driveId: 'D1', itemId: 'I1' },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    expect(parse(r)).toMatchObject({
      attachmentId: 'A1',
      name: 'AG Letter.pdf',
      size: 2000,
      uploadMethod: 'inline',
      draftWebLink: 'https://outlook/d1',
    });
    noBase64(r, buf);
  });

  it('large file: attachment upload session in ranges, id read from Location', async () => {
    const size = 4 * 1024 * 1024;
    const buf = bytes(size, 11);
    const received: Buffer[] = [];
    mockDraft();
    mockDriveFile(buf);
    mockAttachmentUploadSession(size, received);
    const r = await t().execute(
      { draftMessageId: 'DRAFT1', driveId: 'D1', itemId: 'I1', attachmentName: 'Exhibit B.pdf' },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    expect(parse(r)).toMatchObject({
      attachmentId: 'NEWATT=',
      name: 'Exhibit B.pdf',
      size,
      uploadMethod: 'uploadSession',
    });
    const session = calls.find((c) => c.url.endsWith('/attachments/createUploadSession'))!;
    expect(JSON.parse(session.body as string)).toEqual({
      AttachmentItem: {
        attachmentType: 'file',
        name: 'Exhibit B.pdf',
        size,
        contentType: 'application/pdf',
      },
    });
    expect(Buffer.concat(received).equals(buf)).toBe(true);
    for (const p of calls.filter((c) => c.url === UPLOAD_URL)) {
      expect(p.headers.authorization).toBeUndefined();
      expect(Number(p.headers['content-length'])).toBeLessThan(4 * 1024 * 1024);
    }
    expect(
      calls.some((c) => c.method === 'POST' && c.url === `${G}/me/messages/DRAFT1/attachments`)
    ).toBe(false);
    noBase64(r, buf);
  });

  it('refuses a message that is not a draft, a folder, a missing item, and over 150 MB', async () => {
    mockDraft(false);
    const notDraft = await t().execute(
      { draftMessageId: 'DRAFT1', driveId: 'D1', itemId: 'I1' },
      makeCtx()
    );
    expect(parse(notDraft).error).toMatch(/is not a draft/);

    handlers = [];
    mockDraft();
    on('GET', ITEM_META, () => json({ id: 'I1', name: 'Folder', folder: { childCount: 2 } }));
    const folder = await t().execute(
      { draftMessageId: 'DRAFT1', driveId: 'D1', itemId: 'I1' },
      makeCtx()
    );
    expect(parse(folder).error).toMatch(/is a folder/);

    handlers = [];
    mockDraft();
    on('GET', ITEM_META, () => graphError(404, 'itemNotFound'));
    const missing = await t().execute(
      { draftMessageId: 'DRAFT1', driveId: 'D1', itemId: 'I1' },
      makeCtx()
    );
    expect(parse(missing).error).toMatch(/Drive item I1 in drive D1 was not found/);

    handlers = [];
    calls = [];
    mockDraft();
    mockDriveFile(bytes(10), 'huge.pdf', FILE_OPS_MAX_BYTES + 1);
    const huge = await t().execute(
      { draftMessageId: 'DRAFT1', driveId: 'D1', itemId: 'I1' },
      makeCtx()
    );
    expect(parse(huge).error).toMatch(/over the 150 MB/);
    expect(calls.some((c) => c.url === ITEM_CONTENT)).toBe(false);
  });
});

describe('attach_mail_attachment_to_draft', () => {
  const t = () => tool('attach_mail_attachment_to_draft');

  it('small attachment: inline POST', async () => {
    const buf = bytes(500);
    mockDraft();
    mockAttachment(buf, 'Signed Exhibit A.pdf');
    on('POST', `${G}/me/messages/DRAFT1/attachments`, (call) => {
      const body = JSON.parse(call.body as string);
      expect(Buffer.from(body.contentBytes, 'base64').equals(buf)).toBe(true);
      return json({ id: 'A2', name: body.name }, 201);
    });
    const r = await t().execute(
      { sourceMessageId: 'MSG1', attachmentId: 'ATT1', draftMessageId: 'DRAFT1' },
      makeCtx()
    );
    expect(parse(r)).toMatchObject({
      attachmentId: 'A2',
      name: 'Signed Exhibit A.pdf',
      size: 500,
      uploadMethod: 'inline',
    });
    noBase64(r, buf);
    expect(logLines.some((l) => l.includes('"tool":"attach_mail_attachment_to_draft"'))).toBe(true);
  });

  it('large attachment: upload session', async () => {
    const size = 3 * 1024 * 1024 + 17;
    const buf = bytes(size, 5);
    const received: Buffer[] = [];
    mockDraft();
    mockAttachment(buf, 'Scan.pdf');
    mockAttachmentUploadSession(size, received);
    const r = await t().execute(
      { sourceMessageId: 'MSG1', attachmentId: 'ATT1', draftMessageId: 'DRAFT1' },
      makeCtx()
    );
    expect(parse(r)).toMatchObject({
      attachmentId: 'NEWATT=',
      size,
      uploadMethod: 'uploadSession',
    });
    expect(Buffer.concat(received).equals(buf)).toBe(true);
  });

  it('falls back to the attachment list when the Location header is absent', async () => {
    const size = 3 * 1024 * 1024;
    const buf = bytes(size);
    mockDraft();
    mockAttachment(buf, 'Scan.pdf');
    on('POST', `${G}/me/messages/DRAFT1/attachments/createUploadSession`, () =>
      json({ uploadUrl: UPLOAD_URL })
    );
    let sent = 0;
    on('PUT', UPLOAD_URL, (call) => {
      sent += (call.body as Buffer).length;
      return sent < size ? json({}, 200) : new Response(null, { status: 201 });
    });
    on('GET', `${G}/me/messages/DRAFT1/attachments?$select=id,name,size`, () =>
      json({
        value: [
          { id: 'OLD', name: 'Other.pdf' },
          { id: 'LISTED', name: 'Scan.pdf' },
        ],
      })
    );
    const r = await t().execute(
      { sourceMessageId: 'MSG1', attachmentId: 'ATT1', draftMessageId: 'DRAFT1' },
      makeCtx()
    );
    expect(parse(r).attachmentId).toBe('LISTED');
  });
});

describe('find_matter_folder', () => {
  const t = () => tool('find_matter_folder');

  function mockTree(driveId = DEFAULT_MATTERS_DRIVE_ID, rootPath = 'GC%20Associates/Matters') {
    const d = encodeURIComponent(driveId);
    on('GET', `${G}/drives/${d}/root:/${rootPath}?$select=id,name,webUrl,folder`, () =>
      json({ id: 'ROOT', name: 'Matters', folder: {} })
    );
    const children = (id: string) =>
      `${G}/drives/${d}/items/${id}/children?$select=id,name,webUrl,folder&$top=999`;
    on('GET', children('ROOT'), () =>
      json({
        value: [
          { id: 'C1', name: 'BODEC', webUrl: 'https://sp/BODEC', folder: {} },
          { id: 'FILE', name: 'bodec notes.docx', file: {} },
        ],
        '@odata.nextLink': `${G}/drives/${d}/items/ROOT/children?$skiptoken=p2`,
      })
    );
    on('GET', `${G}/drives/${d}/items/ROOT/children?$skiptoken=p2`, () =>
      json({
        value: [
          { id: 'C2', name: 'Buckner HeavyLift', webUrl: 'https://sp/Buckner', folder: {} },
          { id: 'C3', name: 'Other', webUrl: 'https://sp/Other', folder: {} },
        ],
      })
    );
    on('GET', children('C1'), () =>
      json({ value: [{ id: 'M1', name: 'NBU Public Information Request', folder: {} }] })
    );
    on('GET', children('C2'), () =>
      json({
        value: [{ id: 'M2', name: 'Bodec Subcontract', webUrl: 'https://sp/bs', folder: {} }],
      })
    );
    on('GET', children('C3'), () => json({ value: [] }));
  }

  it('matches client and matter folders case-insensitively, exact first, files ignored', async () => {
    mockTree();
    const r = await t().execute({ query: 'Bodec' }, makeCtx());
    expect(r.isError).toBeFalsy();
    const out = parse(r);
    expect(out).toEqual([
      {
        client: 'BODEC',
        matter: null,
        itemId: 'C1',
        driveId: DEFAULT_MATTERS_DRIVE_ID,
        path: 'GC Associates/Matters/BODEC',
        webUrl: 'https://sp/BODEC',
        exact: true,
      },
      {
        client: 'Buckner HeavyLift',
        matter: 'Bodec Subcontract',
        itemId: 'M2',
        driveId: DEFAULT_MATTERS_DRIVE_ID,
        path: 'GC Associates/Matters/Buckner HeavyLift/Bodec Subcontract',
        webUrl: 'https://sp/bs',
        exact: false,
      },
    ]);
    expect(logLines.some((l) => l.includes('"tool":"find_matter_folder"'))).toBe(true);
  });

  it('finds a matter by part of its name, honors maxResults, and resolves the root once', async () => {
    mockTree();
    const r1 = await t().execute({ query: 'public information' }, makeCtx());
    expect(parse(r1)).toEqual([
      expect.objectContaining({
        client: 'BODEC',
        matter: 'NBU Public Information Request',
        itemId: 'M1',
      }),
    ]);
    const r2 = await t().execute({ query: 'b', maxResults: 1 }, makeCtx());
    expect(parse(r2)).toHaveLength(1);
    expect(calls.filter((c) => c.url.includes('/root:/'))).toHaveLength(1);
    const none = await t().execute({ query: 'zzz' }, makeCtx());
    expect(parse(none)).toEqual([]);
  });

  it('uses MATTERS_DRIVE_ID and MATTERS_ROOT_PATH when set', async () => {
    process.env.MATTERS_DRIVE_ID = 'b!OLDRIVE';
    process.env.MATTERS_ROOT_PATH = '/OL Docs/Clients/';
    mockTree('b!OLDRIVE', 'OL%20Docs/Clients');
    const r = await t().execute({ query: 'other' }, makeCtx());
    expect(parse(r)).toEqual([
      expect.objectContaining({
        client: 'Other',
        driveId: 'b!OLDRIVE',
        path: 'OL Docs/Clients/Other',
      }),
    ]);
  });

  it('refuses an empty query and reports a missing root clearly', async () => {
    expect(parse(await t().execute({ query: '  ' }, makeCtx())).error).toMatch(/query is required/);
    const r = await t().execute({ query: 'x' }, makeCtx());
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/Matters root "GC Associates\/Matters" .* was not found/);
  });
});
