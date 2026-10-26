import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { GraphClient } from '../src/graph-client.js';

const { info, debug } = vi.hoisted(() => ({ info: vi.fn(), debug: vi.fn() }));
vi.mock('../src/logger.js', () => ({
  default: { info, debug, error: vi.fn(), warn: vi.fn() },
}));

import { __resetRegistrationLogStateForTests, registerGraphTools } from '../src/graph-tools.js';

// HTTP/OAuth mode re-registers every tool on every request. Before this change each
// registration logged the whole --enabled-tools pattern plus one "Skipping tool" line per
// skipped tool (315 lines on the firm preview), which overran Railway's per-second log limit
// and dropped the audit lines.
describe('registration logging', () => {
  beforeEach(() => {
    info.mockClear();
    debug.mockClear();
    __resetRegistrationLogStateForTests();
  });

  function register(pattern: string) {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    return registerGraphTools(server, {} as GraphClient, {
      orgMode: true,
      enabledTools: pattern,
    });
  }

  const infoLines = () => info.mock.calls.map((c) => String(c[0]));

  it('logs one summary line at info per registration and the skips at debug', () => {
    const registered = register('^(list-mail-messages|find-matter-folder)$');
    expect(registered).toBe(2);
    const lines = infoLines();
    expect(lines.filter((l) => l.startsWith('Skipping'))).toEqual([]);
    expect(lines.filter((l) => l.startsWith('Tool registration complete'))).toHaveLength(1);
    expect(lines.filter((l) => l.startsWith('Tool filtering enabled'))).toHaveLength(1);
    expect(lines.length).toBeLessThan(10);
    expect(debug.mock.calls.some((c) => String(c[0]).startsWith('Skipping tool'))).toBe(true);
  });

  it('echoes the filter pattern at info only once per process', () => {
    register('^(list-mail-messages)$');
    info.mockClear();
    register('^(list-mail-messages)$');
    const lines = infoLines();
    expect(lines.filter((l) => l.startsWith('Tool filtering enabled'))).toEqual([]);
    expect(lines).toEqual([expect.stringMatching(/^Tool registration complete: \d+ registered/)]);
    expect(debug.mock.calls.some((c) => String(c[0]).startsWith('Tool filtering enabled'))).toBe(
      true
    );
  });
});
