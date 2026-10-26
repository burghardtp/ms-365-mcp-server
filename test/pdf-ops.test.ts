import { beforeEach, afterEach, describe, expect, it, vi } from 'vitest';
import { PDFDocument, StandardFonts } from 'pdf-lib';

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
  firstTextLine,
  parsePageRange,
  pdfFirstLines,
  PDF_INFO_LINE_MAX_CHARS,
  PDF_MAX_INPUT_BYTES,
} from '../src/pdf-ops.js';
import { __resetSharedBreakerForTests } from '../src/lib/graph-resilience.js';
import { TOOL_CATEGORIES } from '../src/tool-categories.js';

const G = 'https://graph.microsoft.com/v1.0';

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

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

function binary(buf: Buffer | Uint8Array): Response {
  return new Response(buf, {
    status: 200,
    headers: { 'content-type': 'application/pdf', 'content-length': String(buf.length) },
  });
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

/** A PDF with one page per entry; each page shows its lines of text (an empty entry is a blank page). */
async function makePdf(pages: string[], size = 12): Promise<Buffer> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (const text of pages) {
    const page = doc.addPage([612, 792]);
    let y = 700;
    for (const line of text.split('\n')) {
      if (line.trim()) page.drawText(line, { x: 20, y, size, font });
      y -= 18;
    }
  }
  return Buffer.from(await doc.save());
}

async function firstLinesOf(buf: Buffer): Promise<string[]> {
  return (await pdfFirstLines(buf)).firstLines;
}

function mockDriveFile(
  driveId: string,
  itemId: string,
  buf: Buffer,
  name = `${itemId}.pdf`,
  declared = buf.length
) {
  on('GET', `${G}/drives/${driveId}/items/${itemId}?$select=id,name,size,webUrl,file,folder`, () =>
    json({ id: itemId, name, size: declared, file: { mimeType: 'application/pdf' } })
  );
  on('GET', `${G}/drives/${driveId}/items/${itemId}/content`, () => binary(buf));
}

/** Capture the single-PUT upload into a folder and answer with a new drive item. */
function mockUpload(driveId: string, folderId: string, returnedName?: string) {
  const uploads: Call[] = [];
  on('PUT', new RegExp(`^${G}/drives/${driveId}/items/${folderId}:/[^/]+:/content\\?`), (call) => {
    uploads.push(call);
    const name = decodeURIComponent(
      (call.url.match(/\/items\/[^:]+:\/([^:]+):\/content/) ?? [])[1] ?? ''
    );
    const body = call.body as Buffer;
    return json({
      id: 'NEWITEM',
      name: returnedName ?? name,
      size: body.length,
      webUrl: `https://sp/${encodeURIComponent(returnedName ?? name)}`,
    });
  });
  return uploads;
}

beforeEach(() => {
  calls = [];
  handlers = [];
  logLines.length = 0;
  __resetSharedBreakerForTests();
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
      return json(
        { error: { code: 'unmocked', message: `unmocked ${call.method} ${call.url}` } },
        404
      );
    })
  );
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('registration', () => {
  it('registers the three PDF tools in the file presets', () => {
    for (const name of ['pdf-merge', 'pdf-replace-pages', 'pdf-info']) {
      expect(tool(name)).toBeDefined();
      expect(TOOL_CATEGORIES.files.pattern.test(name)).toBe(true);
      expect(TOOL_CATEGORIES.onedrive.pattern.test(name)).toBe(true);
    }
    expect(tool('pdf-info').readOnlyHint).toBe(true);
    expect(tool('pdf-merge').readOnlyHint).toBe(false);
    expect(tool('pdf-replace-pages').readOnlyHint).toBe(false);
  });
});

describe('helpers', () => {
  it('parses page ranges and refuses bad ones', () => {
    expect(parsePageRange('3', 5)).toEqual({ from: 3, to: 3 });
    expect(parsePageRange(' 2 - 4 ', 5)).toEqual({ from: 2, to: 4 });
    expect(parsePageRange('0', 5)).toMatch(/not a valid range/);
    expect(parsePageRange('4-2', 5)).toMatch(/not a valid range/);
    expect(parsePageRange('5-6', 5)).toMatch(/past the end .*5 pages/);
    expect(parsePageRange('two', 5)).toMatch(/must be "a" or "a-b"/);
  });

  it('takes the first non-empty line and caps it', () => {
    expect(
      firstTextLine([{ str: '  ', hasEOL: true }, { str: 'EXHIBIT', hasEOL: false }, { str: ' A' }])
    ).toBe('EXHIBIT A');
    expect(firstTextLine([])).toBe('');
    expect(firstTextLine([{ str: 'x'.repeat(500) }])).toHaveLength(PDF_INFO_LINE_MAX_CHARS);
  });
});

describe('pdf-info', () => {
  it('returns the page count and each page first line, "" for a blank page', async () => {
    const buf = await makePdf(['EXHIBIT A\nDeclaration of Anita', '', 'EXHIBIT B\nLetter']);
    mockDriveFile('D1', 'I1', buf, 'Exhibits.pdf');
    const r = await tool('pdf-info').execute({ driveId: 'D1', itemId: 'I1' }, makeCtx());
    expect(r.isError).toBeFalsy();
    expect(parse(r)).toEqual({
      name: 'Exhibits.pdf',
      size: buf.length,
      pageCount: 3,
      pages: [
        { page: 1, firstLine: 'EXHIBIT A' },
        { page: 2, firstLine: '' },
        { page: 3, firstLine: 'EXHIBIT B' },
      ],
    });
    const lines = logLines.filter((l) => l.includes('"tool":"pdf-info"'));
    expect(lines).toHaveLength(1);
    expect(lines[0]).toContain('"destination":"pages:3"');
  });

  it('caps a long first line at 200 characters', async () => {
    const long = 'A'.repeat(120) + ' ' + 'B'.repeat(120);
    const buf = await makePdf([long], 3);
    mockDriveFile('D1', 'I1', buf);
    const out = parse(await tool('pdf-info').execute({ driveId: 'D1', itemId: 'I1' }, makeCtx()));
    expect(out.pages[0].firstLine).toHaveLength(200);
    expect(out.pages[0].firstLine.startsWith('AAAA')).toBe(true);
  });

  it('refuses a file over 100 MB before downloading it', async () => {
    mockDriveFile('D1', 'BIG', Buffer.from('%PDF-'), 'big.pdf', PDF_MAX_INPUT_BYTES + 1);
    const r = await tool('pdf-info').execute({ driveId: 'D1', itemId: 'BIG' }, makeCtx());
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/over the 100 MB/);
    expect(calls.some((c) => c.url.endsWith('/content'))).toBe(false);
  });

  it('refuses a file that is not a PDF', async () => {
    mockDriveFile('D1', 'DOC', Buffer.from('PK\u0003\u0004 not a pdf'), 'memo.docx');
    const r = await tool('pdf-info').execute({ driveId: 'D1', itemId: 'DOC' }, makeCtx());
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/"memo.docx" is not a PDF/);
  });
});

describe('pdf-merge', () => {
  it('merges the parts in order into a new file with rename on conflict', async () => {
    mockDriveFile('D1', 'P1', await makePdf(['Brief page 1', 'Brief page 2']));
    mockDriveFile('D2', 'P2', await makePdf(['EXHIBIT A', 'A2', 'A3']));
    const uploads = mockUpload('OUT', 'FOLDER', 'Brief with Exhibits 1.pdf');
    const r = await tool('pdf-merge').execute(
      {
        parts: [
          { driveId: 'D1', itemId: 'P1' },
          { driveId: 'D2', itemId: 'P2' },
        ],
        outputDriveId: 'OUT',
        outputFolderItemId: 'FOLDER',
        outputFileName: 'Brief with Exhibits: final',
      },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    const out = parse(r);
    expect(out).toMatchObject({
      itemId: 'NEWITEM',
      name: 'Brief with Exhibits 1.pdf',
      pageCount: 5,
      webUrl: expect.stringContaining('https://sp/'),
    });
    expect(uploads).toHaveLength(1);
    expect(uploads[0].url).toContain(':/Brief%20with%20Exhibits%20final.pdf:/content');
    expect(uploads[0].url).toContain('conflictBehavior=rename');
    expect(out.size).toBe((uploads[0].body as Buffer).length);
    expect(await firstLinesOf(uploads[0].body as Buffer)).toEqual([
      'Brief page 1',
      'Brief page 2',
      'EXHIBIT A',
      'A2',
      'A3',
    ]);
    expect(logLines.filter((l) => l.includes('"tool":"pdf-merge"'))).toHaveLength(1);
    expect(JSON.stringify(out)).not.toContain('JVBER'); // no base64 PDF in the result
  });

  it('refuses fewer than two parts, a mail part, and a missing name, with no Graph call', async () => {
    const base = { outputDriveId: 'OUT', outputFolderItemId: 'F', outputFileName: 'x.pdf' };
    const one = await tool('pdf-merge').execute(
      { ...base, parts: [{ driveId: 'D', itemId: 'I' }] },
      makeCtx()
    );
    expect(parse(one).error).toMatch(/at least two/);
    const mail = await tool('pdf-merge').execute(
      {
        ...base,
        parts: [
          { driveId: 'D', itemId: 'I' },
          { messageId: 'M', attachmentId: 'A' },
        ],
      },
      makeCtx()
    );
    expect(parse(mail).error).toMatch(/parts\[1\] must be \{ driveId, itemId \}/);
    const noName = await tool('pdf-merge').execute(
      {
        ...base,
        outputFileName: ' ? ',
        parts: [
          { driveId: 'D', itemId: 'I' },
          { driveId: 'D', itemId: 'J' },
        ],
      },
      makeCtx()
    );
    expect(parse(noName).error).toMatch(/outputFileName is required/);
    expect(calls).toHaveLength(0);
  });

  it('refuses an oversized part and writes nothing', async () => {
    mockDriveFile('D1', 'P1', await makePdf(['one']));
    mockDriveFile('D1', 'BIG', Buffer.from('%PDF-'), 'big.pdf', PDF_MAX_INPUT_BYTES + 1);
    const uploads = mockUpload('OUT', 'F');
    const r = await tool('pdf-merge').execute(
      {
        parts: [
          { driveId: 'D1', itemId: 'P1' },
          { driveId: 'D1', itemId: 'BIG' },
        ],
        outputDriveId: 'OUT',
        outputFolderItemId: 'F',
        outputFileName: 'out.pdf',
      },
      makeCtx()
    );
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/100 MB/);
    expect(uploads).toHaveLength(0);
    expect(logLines.some((l) => l.includes('"tool":"pdf-merge"') && l.includes('FAILED'))).toBe(
      true
    );
  });
});

describe('pdf-replace-pages', () => {
  function mockBase(buf: Buffer) {
    on('GET', `${G}/drives/D1/items/BASE?$select=id,name,parentReference`, () =>
      json({ id: 'BASE', name: 'Packet.pdf', parentReference: { id: 'PARENT', driveId: 'D1' } })
    );
    mockDriveFile('D1', 'BASE', buf, 'Packet.pdf');
  }

  it('swaps the page range for the replacement pages and saves a new file beside the base', async () => {
    mockBase(await makePdf(['Page 1', 'Page 2', 'Page 3', 'Page 4', 'Page 5']));
    on('GET', `${G}/me/messages/MSG/attachments/ATT?$select=id,name,contentType,size`, () =>
      json({
        '@odata.type': '#microsoft.graph.fileAttachment',
        id: 'ATT',
        name: 'Signed.pdf',
        size: 10,
      })
    );
    const signed = await makePdf(['SIGNED EXHIBIT A', 'SIGNED EXHIBIT A p2']);
    on('GET', `${G}/me/messages/MSG/attachments/ATT/$value`, () => binary(signed));
    const uploads = mockUpload('D1', 'PARENT');
    const r = await tool('pdf-replace-pages').execute(
      {
        driveId: 'D1',
        itemId: 'BASE',
        removePages: '2-3',
        replacement: { messageId: 'MSG', attachmentId: 'ATT' },
        outputFileName: 'Packet (signed)',
      },
      makeCtx()
    );
    expect(r.isError).toBeFalsy();
    const out = parse(r);
    expect(out).toMatchObject({
      itemId: 'NEWITEM',
      name: 'Packet (signed).pdf',
      pageCount: 5,
      removed: { from: 2, to: 3, count: 2 },
      inserted: { at: 2, count: 2 },
    });
    expect(uploads).toHaveLength(1);
    expect(uploads[0].url).toContain('/items/PARENT:/Packet%20(signed).pdf:/content');
    expect(uploads[0].url).toContain('conflictBehavior=rename');
    expect(await firstLinesOf(uploads[0].body as Buffer)).toEqual([
      'Page 1',
      'SIGNED EXHIBIT A',
      'SIGNED EXHIBIT A p2',
      'Page 4',
      'Page 5',
    ]);
    // The base item is never written.
    expect(calls.some((c) => c.method !== 'GET' && c.url.includes('/items/BASE'))).toBe(false);
    expect(logLines.filter((l) => l.includes('"tool":"pdf-replace-pages"'))).toHaveLength(1);
  });

  it('takes a drive replacement and a single page', async () => {
    mockBase(await makePdf(['Page 1', 'Page 2', 'Page 3']));
    mockDriveFile('D9', 'NEW', await makePdf(['Replacement']));
    const uploads = mockUpload('D1', 'PARENT');
    const out = parse(
      await tool('pdf-replace-pages').execute(
        {
          driveId: 'D1',
          itemId: 'BASE',
          removePages: '3',
          replacement: { driveId: 'D9', itemId: 'NEW' },
          outputFileName: 'Packet.pdf',
        },
        makeCtx()
      )
    );
    expect(out.removed).toEqual({ from: 3, to: 3, count: 1 });
    expect(await firstLinesOf(uploads[0].body as Buffer)).toEqual([
      'Page 1',
      'Page 2',
      'Replacement',
    ]);
    // Same name as the base: still a new item through rename, never a replace.
    expect(uploads[0].url).toContain('conflictBehavior=rename');
    expect(uploads[0].url).not.toContain('/items/BASE');
  });

  it('refuses a range past the end and writes nothing', async () => {
    mockBase(await makePdf(['Page 1', 'Page 2']));
    mockDriveFile('D9', 'NEW', await makePdf(['Replacement']));
    const uploads = mockUpload('D1', 'PARENT');
    const r = await tool('pdf-replace-pages').execute(
      {
        driveId: 'D1',
        itemId: 'BASE',
        removePages: '2-4',
        replacement: { driveId: 'D9', itemId: 'NEW' },
        outputFileName: 'x',
      },
      makeCtx()
    );
    expect(r.isError).toBe(true);
    expect(parse(r).error).toMatch(/past the end of the document \(2 pages\)/);
    expect(uploads).toHaveLength(0);
  });

  it('refuses a malformed replacement before any Graph call', async () => {
    const r = await tool('pdf-replace-pages').execute(
      {
        driveId: 'D1',
        itemId: 'BASE',
        removePages: '1',
        replacement: { driveId: 'D9', attachmentId: 'A' },
        outputFileName: 'x',
      },
      makeCtx()
    );
    expect(parse(r).error).toMatch(/replacement must be/);
    expect(calls).toHaveLength(0);
  });
});
