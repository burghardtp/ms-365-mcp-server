import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../src/graph-client.js';
import { registerGraphTools } from '../src/graph-tools.js';
import {
  ATTACHMENT_MISSING_MESSAGE,
  GATE_NOT_CONFIGURED_MESSAGE,
  bodyText,
  buildTarget,
  factsFromParams,
  newTextOfBody,
} from '../src/send-gate.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

import logger from '../src/logger.js';

// The seven tools the firm design gates (section E1), marked sendGate in endpoints.json.
const SEND_GATED_TOOLS = new Set([
  'send-mail',
  'send-draft-message',
  'reply-mail-message',
  'reply-all-mail-message',
  'forward-mail-message',
  'send-shared-mailbox-mail',
  'share-drive-item',
]);

const GATE_URL = 'https://context.example.test/gate/check';
const TOKEN = 'unit-test-token-not-a-secret';

type Handler = (params: Record<string, unknown>) => Promise<{
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
}>;

function parse(result: { content: Array<{ text: string }> }, index = 0): Record<string, unknown> {
  return JSON.parse(result.content[index].text);
}

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('server-side send gate', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn>; registerTool: ReturnType<typeof vi.fn> };
  let graphRequest: ReturnType<typeof vi.fn>;
  let makeRequest: ReturnType<typeof vi.fn>;
  let fetchMock: ReturnType<typeof vi.fn>;
  let graphClient: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubEnv('GATE_DISABLED', '');
    vi.stubEnv('GATE_CHECK_URL', GATE_URL);
    vi.stubEnv('SEND_GATE_TOKEN', TOKEN);
    vi.stubEnv('GATE_CONNECTOR_NAME', 'GC365-Unit');
    vi.stubEnv('MS365_MCP_REQUIRE_CONFIRM', '');
    mockServer = { tool: vi.fn(), registerTool: vi.fn() };
    graphRequest = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{"ok":true}' }] });
    makeRequest = vi.fn();
    graphClient = { graphRequest, makeRequest } as unknown as GraphClient;
    fetchMock = vi.fn();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllEnvs();
    vi.unstubAllGlobals();
  });

  function handler(name: string): Handler {
    registerGraphTools(mockServer as never, graphClient, { orgMode: true });
    const call = mockServer.registerTool.mock.calls.find((c: unknown[]) => c[0] === name);
    expect(call).toBeDefined();
    return call![call!.length - 1] as Handler;
  }

  function draftWith(body: string, attachments: unknown[]) {
    makeRequest.mockImplementation(async (endpoint: string) => {
      if (endpoint.includes('/attachments?')) return { value: attachments };
      return {
        subject: 'Tender of Defense',
        toRecipients: [{ emailAddress: { address: 'counsel@example.com', name: 'Counsel' } }],
        ccRecipients: [{ emailAddress: { address: 'client@example.com' } }],
        bccRecipients: [],
        body: { contentType: 'html', content: body },
        hasAttachments: attachments.length > 0,
      };
    });
  }

  it('adds a required gateRecordId to all seven send and share tools, and to no other tool', () => {
    registerGraphTools(mockServer as never, graphClient, { orgMode: true });
    const calls = mockServer.registerTool.mock.calls as Array<
      [
        string,
        { inputSchema: { shape: Record<string, { isOptional(): boolean; description?: string }> } },
      ]
    >;
    const gated = calls
      .filter(([, def]) => 'gateRecordId' in def.inputSchema.shape)
      .map(([n]) => n);
    expect(gated.sort()).toEqual([...SEND_GATED_TOOLS].sort());
    for (const [name, def] of calls) {
      if (!SEND_GATED_TOOLS.has(name)) continue;
      const field = def.inputSchema.shape.gateRecordId;
      expect(field.isOptional()).toBe(false);
      expect(field.description).toContain('send_gate_record on Paul Context');
    }
  });

  it('fails closed: with GATE_CHECK_URL unset every send is refused and nothing is called', async () => {
    vi.stubEnv('GATE_CHECK_URL', '');
    const result = await handler('send-mail')({
      gateRecordId: 'sg-1',
      body: {
        message: { subject: 'Hi', toRecipients: [{ emailAddress: { address: 'a@x.com' } }] },
      },
    });
    expect(result.isError).toBe(true);
    expect(parse(result).reason).toBe(GATE_NOT_CONFIGURED_MESSAGE);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(graphRequest).not.toHaveBeenCalled();
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('GATE_DISABLED=true lets a send through only when GATE_CHECK_URL is unset', async () => {
    vi.stubEnv('GATE_CHECK_URL', '');
    vi.stubEnv('GATE_DISABLED', 'true');
    const result = await handler('reply-mail-message')({
      messageId: 'AAMk1',
      body: { comment: 'Thanks' },
    });
    expect(result.isError).toBeFalsy();
    expect(fetchMock).not.toHaveBeenCalled();
    expect(graphRequest).toHaveBeenCalledTimes(1);
  });

  it('refuses a send without a gateRecordId and never asks the gate or Graph', async () => {
    const result = await handler('send-draft-message')({ messageId: 'AAMk1' });
    expect(result.isError).toBe(true);
    expect(String(parse(result).reason)).toContain('No gateRecordId was given');
    expect(fetchMock).not.toHaveBeenCalled();
    expect(graphRequest).not.toHaveBeenCalled();
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('refuses when the gate answers ok:false, with the reason, and makes no Graph send', async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({
        ok: false,
        reason: 'Send gate record sg-9 does not exist (or expired and was swept).',
      })
    );
    const result = await handler('forward-mail-message')({
      gateRecordId: 'sg-9',
      messageId: 'AAMk2',
      body: { comment: 'FYI', toRecipients: [{ emailAddress: { address: 'b@y.com' } }] },
    });
    expect(result.isError).toBe(true);
    const payload = parse(result);
    expect(payload.reason).toBe('Send gate record sg-9 does not exist (or expired and was swept).');
    expect(payload.gateRecordId).toBe('sg-9');
    expect(graphRequest).not.toHaveBeenCalled();
    const [url, init] = fetchMock.mock.calls[0] as [
      string,
      { headers: Record<string, string>; body: string },
    ];
    expect(url).toBe(GATE_URL);
    expect((init.headers as Record<string, string>).Authorization).toBe(`Bearer ${TOKEN}`);
    const sent = JSON.parse(String(init.body));
    expect(sent).toMatchObject({
      id: 'sg-9',
      tool: 'forward-mail-message',
      connector: 'GC365-Unit',
    });
    expect(sent.target).toContain('to: b@y.com');
    expect(sent.target).toContain('message AAMk2');
  });

  it('refuses on an HTTP error from the gate (403, network failure)', async () => {
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ ok: false, reason: 'Invalid send gate token.' }, 403)
    );
    const first = await handler('share-drive-item')({
      gateRecordId: 'sg-3',
      driveId: 'b!drive',
      driveItemId: '01ITEM',
      body: { recipients: [{ email: 'c@z.com' }], roles: ['read'] },
    });
    expect(first.isError).toBe(true);
    expect(parse(first).reason).toBe('Invalid send gate token.');

    fetchMock.mockRejectedValueOnce(new Error('connect ECONNREFUSED'));
    const second = await handler('send-mail')({
      gateRecordId: 'sg-3',
      body: { message: { toRecipients: [{ emailAddress: { address: 'a@x.com' } }] } },
    });
    expect(second.isError).toBe(true);
    expect(String(parse(second).reason)).toContain('ECONNREFUSED');

    fetchMock.mockResolvedValueOnce(new Response('upstream down', { status: 502 }));
    const third = await handler('send-mail')({ gateRecordId: 'sg-3', body: { message: {} } });
    expect(third.isError).toBe(true);
    expect(parse(third).reason).toBe('Send gate check answered HTTP 502.');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('refuses send-draft-message when the body says attached and only inline images are on the draft', async () => {
    draftWith('<p>Please see the <b>attached</b> letter.</p><img src="cid:logo">', [
      { id: 'a1', name: 'logo.png', size: 2048, isInline: true, contentType: 'image/png' },
    ]);
    const result = await handler('send-draft-message')({
      gateRecordId: 'sg-4',
      messageId: 'AAMk4',
    });
    expect(result.isError).toBe(true);
    const payload = parse(result);
    expect(payload.reason).toBe(ATTACHMENT_MISSING_MESSAGE);
    expect(payload.to).toEqual(['counsel@example.com']);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(graphRequest).not.toHaveBeenCalled();
    const endpoints = makeRequest.mock.calls.map((c: unknown[]) => c[0]);
    expect(endpoints).toEqual([
      '/me/messages/AAMk4?$select=subject,toRecipients,ccRecipients,bccRecipients,body,hasAttachments',
      '/me/messages/AAMk4/attachments?$select=id,name,size,isInline,contentType',
    ]);
    expect(logger.info).toHaveBeenCalledWith(
      expect.stringMatching(
        /send-gate ts=\S+ tool=send-draft-message gate=sg-4 recipients=2 attachments=0 result=refused/
      )
    );
  });

  it('sends a gated draft and echoes recipients, subject, and attachment names and sizes', async () => {
    draftWith('<p>Enclosed is the signed Exhibit A.</p>', [
      {
        id: 'a1',
        name: 'Exhibit A (signed).pdf',
        size: 183422,
        isInline: false,
        contentType: 'application/pdf',
      },
      { id: 'a2', name: 'logo.png', size: 2048, isInline: true, contentType: 'image/png' },
    ]);
    fetchMock.mockResolvedValue(
      jsonResponse({
        ok: true,
        id: 'sg-5',
        matter: 'Bodec NBU',
        doc: 'Tender letter',
        ts: 1761500000000,
      })
    );
    const result = await handler('send-draft-message')({
      gateRecordId: 'sg-5',
      messageId: 'AAMk5',
    });
    expect(result.isError).toBeFalsy();

    expect(graphRequest).toHaveBeenCalledTimes(1);
    const [path, options] = graphRequest.mock.calls[0] as [string, { body?: string }];
    expect(path).toBe('/me/messages/AAMk5/send');
    expect(options.body ?? '').not.toContain('gateRecordId');

    const sent = JSON.parse(
      String(
        (fetchMock.mock.calls[0] as [string, { headers: Record<string, string>; body: string }])[1]
          .body
      )
    );
    expect(sent.target).toContain('to: counsel@example.com');
    expect(sent.target).toContain('cc: client@example.com');
    expect(sent.target).toContain('subject: Tender of Defense');
    expect(sent.target).toContain('attachments: Exhibit A (signed).pdf (183422 bytes)');
    expect(sent.target).not.toContain('logo.png');
    expect(sent.target.length).toBeLessThanOrEqual(300);

    const summary = parse(result, result.content.length - 1);
    expect(summary).toEqual({
      sendGate: { id: 'sg-5', matter: 'Bodec NBU', doc: 'Tender letter' },
      tool: 'send-draft-message',
      to: ['counsel@example.com'],
      cc: ['client@example.com'],
      bcc: [],
      subject: 'Tender of Defense',
      attachments: [{ name: 'Exhibit A (signed).pdf', size: 183422 }],
    });
    expect(logger.info).toHaveBeenCalledWith(
      expect.stringMatching(
        /tool=send-draft-message gate=sg-5 recipients=2 attachments=1 result=sent/
      )
    );
    const logged = (logger.info as ReturnType<typeof vi.fn>).mock.calls
      .map((c) => String(c[0]))
      .join('\n');
    expect(logged).not.toContain(TOKEN);
  });

  it('sends send-mail after an ok gate, without forwarding gateRecordId to Graph', async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({ ok: true, id: 'sg-6', matter: 'M', doc: 'D', ts: 1 })
    );
    const result = await handler('send-mail')({
      gateRecordId: 'sg-6',
      body: {
        message: {
          subject: 'Status',
          body: { contentType: 'text', content: 'No attachment here.' },
          toRecipients: [{ emailAddress: { address: 'a@x.com' } }],
          bccRecipients: [{ emailAddress: { address: 'file@x.com' } }],
        },
      },
    });
    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
    const options = graphRequest.mock.calls[0][1] as { body: string };
    expect(options.body).not.toContain('gateRecordId');
    const summary = parse(result, result.content.length - 1);
    expect(summary).toMatchObject({ to: ['a@x.com'], bcc: ['file@x.com'], subject: 'Status' });
    expect(summary).not.toHaveProperty('attachments');
  });

  const OUTLOOK_REPLY_HEADER =
    '<hr style="display:inline-block;width:98%"><p><b>From:</b> Counsel &lt;counsel@example.com&gt;<br>' +
    '<b>Sent:</b> Monday, October 26, 2026 9:14 AM<br><b>To:</b> Paul Burghardt<br>' +
    '<b>Subject:</b> Tender of Defense</p>';

  it('passes a reply draft whose only "attached" is in the quoted thread', async () => {
    draftWith(
      '<p>Thanks, I will review and respond this week.</p>' +
        OUTLOOK_REPLY_HEADER +
        '<p>Please see the attached tender letter and the enclosed policy.</p>',
      []
    );
    fetchMock.mockResolvedValue(jsonResponse({ ok: true, id: 'sg-10' }));
    const result = await handler('send-draft-message')({
      gateRecordId: 'sg-10',
      messageId: 'AAMk10',
    });
    expect(result.isError).toBeFalsy();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(graphRequest).toHaveBeenCalledTimes(1);
    expect(logger.info).toHaveBeenCalledWith(
      expect.stringMatching(
        /tool=send-draft-message gate=sg-10 recipients=2 attachments=0 result=sent/
      )
    );
  });

  it('refuses a reply draft that says "attached" above the quoted thread with no attachment', async () => {
    draftWith(
      '<p>The signed release is attached.</p>' +
        OUTLOOK_REPLY_HEADER +
        '<p>Can you send the release?</p>',
      []
    );
    const result = await handler('send-draft-message')({
      gateRecordId: 'sg-11',
      messageId: 'AAMk11',
    });
    expect(result.isError).toBe(true);
    expect(parse(result).reason).toBe(ATTACHMENT_MISSING_MESSAGE);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('cuts an HTML draft at the divRplyFwdMsg container before testing', async () => {
    draftWith(
      '<div>Confirmed, Tuesday works.</div><div id="divRplyFwdMsg" dir="ltr"><font>earlier message</font></div>' +
        '<div>The enclosed draft is ready for signature.</div>',
      []
    );
    fetchMock.mockResolvedValue(jsonResponse({ ok: true, id: 'sg-12' }));
    const result = await handler('send-draft-message')({
      gateRecordId: 'sg-12',
      messageId: 'AAMk12',
    });
    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
  });

  it('refuses attachmentCheck "skip" without a reason, before reading the draft', async () => {
    draftWith('<p>The report is attached.</p>', []);
    for (const extra of [
      {},
      { attachmentCheckReason: '   ' },
      { attachmentCheckReason: 'x'.repeat(201) },
    ]) {
      const result = await handler('send-draft-message')({
        gateRecordId: 'sg-13',
        messageId: 'AAMk13',
        attachmentCheck: 'skip',
        ...extra,
      });
      expect(result.isError).toBe(true);
      expect(String(parse(result).reason)).toContain(
        'attachmentCheck "skip" requires attachmentCheckReason'
      );
    }
    expect(makeRequest).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('attachmentCheck "skip" with a reason sends, logs the reason, and puts it in the gate target', async () => {
    draftWith('<p>The report attached to my October 20 email still controls.</p>', []);
    fetchMock.mockResolvedValue(jsonResponse({ ok: true, id: 'sg-14' }));
    const result = await handler('send-draft-message')({
      gateRecordId: 'sg-14',
      messageId: 'AAMk14',
      attachmentCheck: 'skip',
      attachmentCheckReason: 'Refers to the report sent on 10.20, not attached by design.',
    });
    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
    const [path, options] = graphRequest.mock.calls[0] as [string, { body?: string }];
    expect(path).toBe('/me/messages/AAMk14/send');
    expect(options.body ?? '').not.toContain('attachmentCheck');
    const sent = JSON.parse(
      String(
        (fetchMock.mock.calls[0] as [string, { headers: Record<string, string>; body: string }])[1]
          .body
      )
    );
    expect(sent.target).toContain(
      'attachment check skipped: Refers to the report sent on 10.20, not attached by design.'
    );
    expect(sent.target).toContain('message AAMk14');
    expect(logger.info).toHaveBeenCalledWith(
      expect.stringMatching(
        /tool=send-draft-message gate=sg-14 recipients=2 attachments=0 attachmentCheck=skip reason="Refers to the report sent on 10\.20, not attached by design\." result=sent/
      )
    );
  });

  it('puts attachmentCheck and attachmentCheckReason on send-draft-message only, both optional', () => {
    registerGraphTools(mockServer as never, graphClient, { orgMode: true });
    const calls = mockServer.registerTool.mock.calls as Array<
      [
        string,
        {
          inputSchema: {
            shape: Record<
              string,
              {
                isOptional(): boolean;
                description?: string;
                safeParse(v: unknown): { success: boolean };
              }
            >;
          };
        },
      ]
    >;
    const withCheck = calls
      .filter(([, def]) => 'attachmentCheck' in def.inputSchema.shape)
      .map(([n]) => n);
    expect(withCheck).toEqual(['send-draft-message']);
    const shape = calls.find(([n]) => n === 'send-draft-message')![1].inputSchema.shape;
    expect(shape.attachmentCheck.isOptional()).toBe(true);
    expect(shape.attachmentCheckReason.isOptional()).toBe(true);
    expect(shape.attachmentCheck.description).toContain('"skip"');
    expect(shape.attachmentCheckReason.description).toContain('200 characters');
    expect(shape.attachmentCheck.safeParse('skip').success).toBe(true);
    expect(shape.attachmentCheck.safeParse('off').success).toBe(false);
    expect(shape.attachmentCheckReason.safeParse('x'.repeat(201)).success).toBe(false);
  });

  it('a Graph failure after an ok gate is reported, not echoed as sent', async () => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true, id: 'sg-7' }));
    graphRequest.mockResolvedValue({
      content: [{ type: 'text', text: '{"error":"x"}' }],
      isError: true,
    });
    const result = await handler('reply-all-mail-message')({
      gateRecordId: 'sg-7',
      messageId: 'AAMk7',
      body: { comment: 'Agreed' },
    });
    expect(result.isError).toBe(true);
    expect(result.content).toHaveLength(1);
    expect(logger.info).toHaveBeenCalledWith(expect.stringMatching(/result=graph-error/));
  });
});

describe('send gate helpers', () => {
  it('strips HTML, comments, styles, and entities from a body', () => {
    expect(
      bodyText(
        '<style>p{}</style><!-- attached --><p>See&nbsp;the&#32;enclosure &amp; reply</p>',
        'html'
      )
    ).toBe('See the enclosure & reply');
    expect(bodyText('<b>not html</b>', 'text')).toBe('<b>not html</b>');
  });

  it('reads recipients from body, body.message, and flattened params', () => {
    const facts = factsFromParams({
      toRecipients: [{ emailAddress: { address: 'flat@x.com' } }],
      body: {
        message: {
          subject: 'S',
          toRecipients: [{ emailAddress: { address: 'a@x.com' } }, 'flat@x.com'],
          ccRecipients: [{ emailAddress: { address: 'c@x.com' } }],
        },
      },
    });
    expect(facts).toEqual({
      to: ['flat@x.com', 'a@x.com'],
      cc: ['c@x.com'],
      bcc: [],
      subject: 'S',
    });
  });

  it('newTextOfBody cuts at each quoted-history marker', () => {
    expect(
      newTextOfBody('Sounds good.\n\n-----Original Message-----\nFrom: A\nSee attached.', 'text')
    ).toBe('Sounds good.');
    expect(
      newTextOfBody(
        'Sounds good.\nOn Mon, Oct 26, 2026 at 9:14 AM Counsel <c@x.com> wrote:\n> attached',
        'text'
      )
    ).toBe('Sounds good.');
    expect(
      newTextOfBody('Sounds good.\n________________________________\nFrom: A\nattached', 'text')
    ).toBe('Sounds good.');
    expect(
      newTextOfBody('Sounds good.\nFrom: A\nTo: B\nCc: C\nDate: Mon\nSee attached.', 'text')
    ).toBe('Sounds good.');
    // A From: line with no Sent: or Date: within four lines is not a header block.
    expect(newTextOfBody('From: the client file\nNotes\nattached here', 'text')).toBe(
      'From: the client file Notes attached here'
    );
    expect(
      newTextOfBody(
        '<div dir="ltr">Thanks.</div><div class="gmail_quote gmail_quote_container"><div>attached</div></div>',
        'html'
      )
    ).toBe('Thanks.');
    expect(
      newTextOfBody('<p>Thanks.</p><div id="appendonsend"></div><hr><div>enclosed</div>', 'html')
    ).toBe('Thanks.');
    expect(newTextOfBody('<p>No quoted thread.</p><p>Second line.</p>', 'html')).toBe(
      'No quoted thread. Second line.'
    );
  });

  it('caps the gate target at 300 characters', () => {
    const many = Array.from({ length: 40 }, (_, i) => ({
      emailAddress: { address: `person${i}@example.com` },
    }));
    const facts = factsFromParams({ body: { message: { toRecipients: many } } });
    expect(buildTarget('send-mail', {}, facts)).toHaveLength(300);
  });
});

describe('send gate in discovery mode', () => {
  it('get-tool-schema lists gateRecordId as required on a gated tool and not on a draft tool', async () => {
    const { api } = await import('../src/generated/client.js');
    const { describeToolSchema } = await import('../src/lib/tool-schema.js');
    const { readFileSync } = await import('fs');
    const config = JSON.parse(
      readFileSync(new URL('../src/endpoints.json', import.meta.url), 'utf8')
    ) as Array<{ toolName: string; sendGate?: boolean }>;
    const describe = (name: string) => {
      const tool = api.endpoints.find((e) => e.alias === name)!;
      const cfg = config.find((c) => c.toolName === name) as never;
      return describeToolSchema(tool, cfg).parameters.find((p) => p.name === 'gateRecordId');
    };
    expect(describe('send-draft-message')).toMatchObject({ required: true });
    const draftTool = api.endpoints.find((e) => e.alias === 'send-draft-message')!;
    const draftCfg = config.find((c) => c.toolName === 'send-draft-message') as never;
    const draftParams = describeToolSchema(draftTool, draftCfg).parameters;
    expect(draftParams.find((p) => p.name === 'attachmentCheck')).toMatchObject({
      required: false,
    });
    expect(draftParams.find((p) => p.name === 'attachmentCheckReason')).toMatchObject({
      required: false,
    });
    const mailTool = api.endpoints.find((e) => e.alias === 'send-mail')!;
    const mailCfg = config.find((c) => c.toolName === 'send-mail') as never;
    expect(
      describeToolSchema(mailTool, mailCfg).parameters.find((p) => p.name === 'attachmentCheck')
    ).toBeUndefined();
    expect(describe('create-draft-email')).toBeUndefined();
  });
});
