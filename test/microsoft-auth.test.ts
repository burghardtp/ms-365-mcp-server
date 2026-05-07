import { describe, expect, it } from 'vitest';
import { getClientSecretForRedirect } from '../src/lib/microsoft-auth.js';

describe('getClientSecretForRedirect', () => {
  it('omits client secret for localhost loopback redirects', () => {
    expect(getClientSecretForRedirect('http://localhost:52713/callback', 'secret')).toBeUndefined();
  });

  it('omits client secret for 127.0.0.1 loopback redirects', () => {
    expect(getClientSecretForRedirect('http://127.0.0.1:52713/callback', 'secret')).toBeUndefined();
  });

  it('keeps client secret for hosted HTTPS redirects', () => {
    expect(getClientSecretForRedirect('https://claude.ai/api/mcp/auth_callback', 'secret')).toBe(
      'secret'
    );
  });

  it('keeps client secret when redirect URI cannot be parsed', () => {
    expect(getClientSecretForRedirect('not a url', 'secret')).toBe('secret');
  });
});
