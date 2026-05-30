import { describe, expect, it } from 'vitest';
import { getPkceChallenge, isLoopbackRedirectUri } from '../src/lib/oauth-proxy.js';

describe('OAuth proxy helpers', () => {
  it('detects localhost and 127.0.0.1 loopback redirect URIs', () => {
    expect(isLoopbackRedirectUri('http://127.0.0.1:58633/callback/random')).toBe(true);
    expect(isLoopbackRedirectUri('http://localhost:58633/callback/random')).toBe(true);
  });

  it('does not treat hosted redirects as loopback redirects', () => {
    expect(isLoopbackRedirectUri('https://claude.ai/api/mcp/auth_callback')).toBe(false);
    expect(isLoopbackRedirectUri('https://example.com/callback')).toBe(false);
  });

  it('supports S256 and plain PKCE challenges', () => {
    expect(getPkceChallenge('verifier', 'plain')).toBe('verifier');
    expect(getPkceChallenge('verifier', 'S256')).toBe(
      'iMnq5o6zALKXGivsnlom_0F5_WYda32GHkxlV7mq7hQ'
    );
  });

  it('rejects unsupported PKCE methods', () => {
    expect(getPkceChallenge('verifier', 'S512')).toBeUndefined();
  });
});
