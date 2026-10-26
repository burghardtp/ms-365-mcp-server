/**
 * Firm PDF tools: merge PDFs, replace a page range, and list each page's first text
 * line, all inside the server. Inputs are read from OneDrive or SharePoint (or a mail
 * attachment for a replacement), outputs are always written as new drive files with
 * conflictBehavior "rename", so an existing file is never overwritten. No tool here
 * returns file bytes to the caller.
 *
 * Tools:
 *   pdf-merge          parts in order -> one new PDF
 *   pdf-replace-pages  base PDF with a page range swapped for another PDF's pages -> new PDF
 *   pdf-info           page count and the first text line of every page
 *
 * Libraries: pdf-lib (merge, remove and insert pages) and pdfjs-dist (text extraction,
 * loaded on first use of pdf-info). Every call writes one log line.
 */
import { z } from 'zod';
import { PDFDocument } from 'pdf-lib';
import type GraphClient from './graph-client.js';
import type { UtilityTool } from './graph-tools.js';
import {
  describe,
  enc,
  fail,
  FILE_OPS_MAX_BYTES,
  FileOpsError,
  graphJson,
  logOp,
  ok,
  readDriveFile,
  readMailAttachment,
  resolveToken,
  sharePointSafeName,
  str,
  withAccount,
  writeToDrive,
  type SourceFile,
} from './file-ops.js';

const MB = 1024 * 1024;
/** Largest PDF any of these tools will read, per input. */
export const PDF_MAX_INPUT_BYTES = 100 * MB;
/** Most parts pdf-merge accepts in one call. */
export const PDF_MERGE_MAX_PARTS = 100;
/** Longest first-line text pdf-info returns per page. */
export const PDF_INFO_LINE_MAX_CHARS = 200;

type PdfSource = { driveId: string; itemId: string } | { messageId: string; attachmentId: string };

function sourceLabel(src: PdfSource): string {
  return 'driveId' in src
    ? `drive:${src.driveId}/item:${src.itemId}`
    : `message:${src.messageId}/attachment:${src.attachmentId}`;
}

async function readPdfSource(
  graph: GraphClient,
  token: string,
  src: PdfSource
): Promise<SourceFile> {
  return 'driveId' in src
    ? readDriveFile(graph, token, src.driveId, src.itemId, PDF_MAX_INPUT_BYTES)
    : readMailAttachment(graph, token, src.messageId, src.attachmentId, PDF_MAX_INPUT_BYTES);
}

function looksLikePdf(bytes: Buffer): boolean {
  // The header may follow a little leading junk; PDF readers accept it within 1 KB.
  return bytes.subarray(0, 1024).includes('%PDF-');
}

async function loadPdf(file: SourceFile): Promise<PDFDocument> {
  if (!looksLikePdf(file.bytes)) {
    throw new FileOpsError(`"${file.name}" is not a PDF (no %PDF- header).`);
  }
  try {
    return await PDFDocument.load(file.bytes, { updateMetadata: false });
  } catch (error) {
    const message = (error as Error)?.message ?? String(error);
    if (/encrypt/i.test(message)) {
      throw new FileOpsError(
        `"${file.name}" is encrypted or password protected. Remove the protection first.`
      );
    }
    throw new FileOpsError(`"${file.name}" could not be read as a PDF: ${message}`);
  }
}

function pdfName(raw: string | undefined): string | undefined {
  const cleaned = sharePointSafeName(raw ?? '');
  if (!cleaned) return undefined;
  return /\.pdf$/i.test(cleaned) ? cleaned : `${cleaned}.pdf`;
}

async function savePdf(doc: PDFDocument): Promise<Buffer> {
  const bytes = Buffer.from(await doc.save());
  if (bytes.length > FILE_OPS_MAX_BYTES) {
    throw new FileOpsError(
      `The result is ${bytes.length} bytes, over the ${FILE_OPS_MAX_BYTES} byte upload limit. Nothing was written.`
    );
  }
  return bytes;
}

/** Parse "a" or "a-b" (1-based, inclusive). */
export function parsePageRange(
  raw: string,
  pageCount: number
): { from: number; to: number } | string {
  const m = raw.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
  if (!m) return `removePages must be "a" or "a-b" (for example "3" or "3-7"), not "${raw}".`;
  const from = Number(m[1]);
  const to = m[2] === undefined ? from : Number(m[2]);
  if (from < 1 || to < from) return `removePages "${raw}" is not a valid range.`;
  if (to > pageCount) {
    return `removePages "${raw}" runs past the end of the document (${pageCount} pages).`;
  }
  return { from, to };
}

// ---------------------------------------------------------------------------
// Text extraction (pdfjs-dist, loaded on first use)

interface PdfJsTextItem {
  str?: string;
  hasEOL?: boolean;
}

interface PdfJsModule {
  getDocument(params: Record<string, unknown>): {
    promise: Promise<{
      numPages: number;
      getPage(n: number): Promise<{
        getTextContent(): Promise<{ items: PdfJsTextItem[] }>;
        cleanup(): void;
      }>;
    }>;
    destroy(): Promise<void>;
  };
}

let pdfjsPromise: Promise<PdfJsModule> | undefined;

function loadPdfJs(): Promise<PdfJsModule> {
  pdfjsPromise ??= import('pdfjs-dist/legacy/build/pdf.mjs') as unknown as Promise<PdfJsModule>;
  return pdfjsPromise;
}

/** First non-empty line of a page's text items, whitespace collapsed, capped. */
export function firstTextLine(items: PdfJsTextItem[]): string {
  let current = '';
  const lines: string[] = [];
  for (const item of items) {
    current += item.str ?? '';
    if (item.hasEOL) {
      lines.push(current);
      current = '';
    }
  }
  lines.push(current);
  for (const line of lines) {
    const t = line.replace(/\s+/g, ' ').trim();
    if (t) return t.slice(0, PDF_INFO_LINE_MAX_CHARS);
  }
  return '';
}

/** Page count and first text line per page. A page with no text layer (scanned) gives "". */
export async function pdfFirstLines(
  bytes: Buffer
): Promise<{ pageCount: number; firstLines: string[] }> {
  const pdfjs = await loadPdfJs();
  const task = pdfjs.getDocument({
    data: new Uint8Array(bytes),
    // Errors only: pdfjs prints warnings with console.log, which would corrupt stdio MCP.
    verbosity: 0,
    isEvalSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
    stopAtErrors: false,
  });
  try {
    let doc;
    try {
      doc = await task.promise;
    } catch (error) {
      const name = (error as { name?: string })?.name ?? '';
      if (name === 'PasswordException') {
        throw new FileOpsError('The PDF is password protected. Remove the protection first.');
      }
      throw new FileOpsError(
        `The PDF could not be read: ${(error as Error)?.message ?? String(error)}`
      );
    }
    const firstLines: string[] = [];
    for (let n = 1; n <= doc.numPages; n++) {
      try {
        const page = await doc.getPage(n);
        const content = await page.getTextContent();
        firstLines.push(firstTextLine(content.items));
        page.cleanup();
      } catch {
        firstLines.push('');
      }
    }
    return { pageCount: doc.numPages, firstLines };
  } finally {
    await task.destroy().catch(() => undefined);
  }
}

// ---------------------------------------------------------------------------
// Tool definitions

const driveRef = z.object({
  driveId: z.string().describe('Drive id of the PDF.'),
  itemId: z.string().describe('Drive item id of the PDF.'),
});

const mailRef = z.object({
  messageId: z.string().describe('Id of the message that carries the PDF.'),
  attachmentId: z.string().describe('Attachment id of the PDF.'),
});

function asSource(value: unknown): PdfSource | undefined {
  const v = value as Record<string, unknown> | undefined;
  if (!v || typeof v !== 'object') return undefined;
  const driveId = str(v.driveId);
  const itemId = str(v.itemId);
  const messageId = str(v.messageId);
  const attachmentId = str(v.attachmentId);
  if (driveId && itemId && !messageId && !attachmentId) return { driveId, itemId };
  if (messageId && attachmentId && !driveId && !itemId) return { messageId, attachmentId };
  return undefined;
}

export const PDF_OPS_TOOLS: UtilityTool[] = [
  {
    name: 'pdf-merge',
    method: 'PUT',
    path: 'tool:pdf-merge',
    searchKeywords: 'merge pdf combine pdfs join pdf files exhibits into one pdf',
    description:
      'Merge PDFs stored in OneDrive or SharePoint, in the order given, into one new PDF saved in the output folder, server side. Never overwrites: if the name exists the new file is saved under a renamed copy. Each input is limited to 100 MB. Returns { itemId, name, size, webUrl, pageCount }.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        parts: z
          .array(driveRef)
          .describe('PDFs to merge, in order: [{ driveId, itemId }, ...] (2 to 100).'),
        outputDriveId: z.string().describe('Drive id of the output folder.'),
        outputFolderItemId: z.string().describe('Item id of the output folder.'),
        outputFileName: z
          .string()
          .describe(
            'Name of the new PDF (".pdf" is added if missing; illegal characters removed).'
          ),
      }),
    execute: async (params, ctx) => {
      const tool = 'pdf-merge';
      const rawParts = Array.isArray(params.parts) ? params.parts : [];
      const parts = rawParts.map(asSource);
      const outputDriveId = str(params.outputDriveId);
      const outputFolderItemId = str(params.outputFolderItemId);
      const outputFileName = pdfName(str(params.outputFileName));
      if (!outputDriveId || !outputFolderItemId) {
        return fail('outputDriveId and outputFolderItemId are required.');
      }
      if (!outputFileName) return fail('outputFileName is required.');
      if (parts.length < 2) return fail('parts must list at least two PDFs.');
      if (parts.length > PDF_MERGE_MAX_PARTS) {
        return fail(`parts may list at most ${PDF_MERGE_MAX_PARTS} PDFs.`);
      }
      const bad = parts.findIndex((p) => !p || !('driveId' in p));
      if (bad >= 0) return fail(`parts[${bad}] must be { driveId, itemId }.`);
      const sources = parts as Array<{ driveId: string; itemId: string }>;
      const destination = `drive:${outputDriveId}/item:${outputFolderItemId}`;
      const sourceText = `parts:${sources.length}[${sources.map((p) => p.itemId).join(',')}]`;
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        const out = await PDFDocument.create();
        let inputBytes = 0;
        for (const src of sources) {
          const file = await readPdfSource(ctx.graphClient, token, src);
          inputBytes += file.bytes.length;
          const doc = await loadPdf(file);
          const pages = await out.copyPages(doc, doc.getPageIndices());
          for (const page of pages) out.addPage(page);
        }
        const bytes = await savePdf(out);
        let written;
        try {
          written = await writeToDrive(
            ctx.graphClient,
            token,
            outputDriveId,
            { itemId: outputFolderItemId },
            outputFileName,
            bytes,
            'rename'
          );
        } catch (error) {
          throw new FileOpsError(describe(error, `Output folder ${outputFolderItemId}`));
        }
        logOp(tool, sourceText, `${destination}/${written.name ?? outputFileName}`, bytes.length);
        return ok({
          itemId: written.id,
          name: written.name ?? outputFileName,
          size: written.size ?? bytes.length,
          webUrl: written.webUrl,
          pageCount: out.getPageCount(),
          inputBytes,
        });
      } catch (error) {
        logOp(tool, sourceText, `${destination} FAILED`, null);
        return fail(describe(error, 'A part'));
      }
    },
  },
  {
    name: 'pdf-replace-pages',
    method: 'PUT',
    path: 'tool:pdf-replace-pages',
    searchKeywords:
      'replace pages in pdf swap exhibit pages substitute signed page insert pdf pages remove pages',
    description:
      'Remove a page range from a PDF in OneDrive or SharePoint and insert another PDF\'s pages at that position, saving the result as a NEW file in the same folder as the base (the base is never changed or overwritten; an existing name gets a renamed copy). removePages is 1-based and inclusive: "5" or "5-8". The replacement is a drive file { driveId, itemId } or a mail attachment { messageId, attachmentId }. Each input is limited to 100 MB. Returns { itemId, name, size, webUrl, pageCount, removed, inserted }.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        driveId: z.string().describe('Drive id of the base PDF.'),
        itemId: z.string().describe('Drive item id of the base PDF.'),
        removePages: z.string().describe('Pages to remove, 1-based inclusive: "a" or "a-b".'),
        replacement: z
          .union([driveRef, mailRef])
          .describe(
            'PDF whose pages go in at the removed position: { driveId, itemId } or { messageId, attachmentId }.'
          ),
        outputFileName: z
          .string()
          .describe(
            'Name of the new PDF (".pdf" is added if missing; illegal characters removed).'
          ),
      }),
    execute: async (params, ctx) => {
      const tool = 'pdf-replace-pages';
      const driveId = str(params.driveId);
      const itemId = str(params.itemId);
      const removeRaw = str(params.removePages);
      const replacement = asSource(params.replacement);
      const outputFileName = pdfName(str(params.outputFileName));
      if (!driveId || !itemId) return fail('driveId and itemId are required.');
      if (!removeRaw) return fail('removePages is required ("a" or "a-b").');
      if (!replacement) {
        return fail('replacement must be { driveId, itemId } or { messageId, attachmentId }.');
      }
      if (!outputFileName) return fail('outputFileName is required.');
      const sourceText = `drive:${driveId}/item:${itemId} pages:${removeRaw} with:${sourceLabel(replacement)}`;
      let destination = `drive:${driveId}/parent-of:${itemId}`;
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        let parentId: string | undefined;
        try {
          const meta = await graphJson<{ parentReference?: { id?: string; driveId?: string } }>(
            ctx.graphClient,
            token,
            `/drives/${enc(driveId)}/items/${enc(itemId)}?$select=id,name,parentReference`
          );
          parentId = meta.parentReference?.id;
        } catch (error) {
          throw new FileOpsError(describe(error, `Drive item ${itemId} in drive ${driveId}`));
        }
        if (!parentId) throw new FileOpsError(`Could not read the folder that holds ${itemId}.`);
        destination = `drive:${driveId}/item:${parentId}`;
        const baseFile = await readPdfSource(ctx.graphClient, token, { driveId, itemId });
        const base = await loadPdf(baseFile);
        const range = parsePageRange(removeRaw, base.getPageCount());
        if (typeof range === 'string') throw new FileOpsError(range);
        const replFile = await readPdfSource(ctx.graphClient, token, replacement);
        const repl = await loadPdf(replFile);
        const removedCount = range.to - range.from + 1;
        for (let i = range.to - 1; i >= range.from - 1; i--) base.removePage(i);
        const pages = await base.copyPages(repl, repl.getPageIndices());
        pages.forEach((page, i) => base.insertPage(range.from - 1 + i, page));
        const bytes = await savePdf(base);
        let written;
        try {
          written = await writeToDrive(
            ctx.graphClient,
            token,
            driveId,
            { itemId: parentId },
            outputFileName,
            bytes,
            'rename'
          );
        } catch (error) {
          throw new FileOpsError(describe(error, `Folder ${parentId}`));
        }
        logOp(tool, sourceText, `${destination}/${written.name ?? outputFileName}`, bytes.length);
        return ok({
          itemId: written.id,
          name: written.name ?? outputFileName,
          size: written.size ?? bytes.length,
          webUrl: written.webUrl,
          pageCount: base.getPageCount(),
          removed: { from: range.from, to: range.to, count: removedCount },
          inserted: { at: range.from, count: pages.length },
        });
      } catch (error) {
        logOp(tool, sourceText, `${destination} FAILED`, null);
        return fail(describe(error, 'The item'));
      }
    },
  },
  {
    name: 'pdf-info',
    method: 'GET',
    path: 'tool:pdf-info',
    searchKeywords: 'pdf page count pdf pages first line find exhibit boundaries pdf outline',
    description:
      'Read a PDF in OneDrive or SharePoint, server side, and return its page count and the first non-empty text line of every page (up to 200 characters), so exhibit boundaries can be found. A page with no text layer (a scan) returns "". Limit 100 MB. Returns { name, size, pageCount, pages: [{ page, firstLine }] }.',
    readOnlyHint: true,
    openWorldHint: true,
    buildSchema: (ctx) =>
      withAccount(ctx, {
        driveId: z.string().describe('Drive id of the PDF.'),
        itemId: z.string().describe('Drive item id of the PDF.'),
      }),
    execute: async (params, ctx) => {
      const tool = 'pdf-info';
      const driveId = str(params.driveId);
      const itemId = str(params.itemId);
      if (!driveId || !itemId) return fail('driveId and itemId are required.');
      const source = `drive:${driveId}/item:${itemId}`;
      try {
        const token = await resolveToken(ctx, params.account as string | undefined);
        const file = await readPdfSource(ctx.graphClient, token, { driveId, itemId });
        if (!looksLikePdf(file.bytes)) {
          throw new FileOpsError(`"${file.name}" is not a PDF (no %PDF- header).`);
        }
        const { pageCount, firstLines } = await pdfFirstLines(file.bytes);
        logOp(tool, source, `pages:${pageCount}`, file.bytes.length);
        return ok({
          name: file.name,
          size: file.bytes.length,
          pageCount,
          pages: firstLines.map((firstLine, i) => ({ page: i + 1, firstLine })),
        });
      } catch (error) {
        logOp(tool, source, 'FAILED', null);
        return fail(describe(error, 'The item'));
      }
    },
  },
];
