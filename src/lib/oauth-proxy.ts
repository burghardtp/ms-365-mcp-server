import type { Request } from 'express';
import crypto from 'node:crypto';

export function getPublicBaseUrl(req: Request, configuredBaseUrl?: string): string {
  if (configuredBaseUrl) {
    return configuredBaseUrl.replace(/\/$/, '');
  }

  const forwardedProto = req.get('x-forwarded-proto')?.split(',')[0]?.trim();
  const protocol = req.secure ? 'https' : forwardedProto || req.protocol || 'http';
  const host = req.get('host') || 'localhost';

  return new URL(`${protocol}://${host}`).origin;
}

export function isLoopbackRedirectUri(redirectUri: string | null): boolean {
  if (!redirectUri) {
    return false;
  }

  try {
    const parsed = new URL(redirectUri);
    const host = parsed.hostname.toLowerCase();

    return parsed.protocol === 'http:' && (host === 'localhost' || host === '127.0.0.1');
  } catch {
    return false;
  }
}

export function getPkceChallenge(verifier: string, method: string): string | undefined {
  if (method.toUpperCase() === 'PLAIN') {
    return verifier;
  }

  if (method.toUpperCase() !== 'S256') {
    return undefined;
  }

  return crypto.createHash('sha256').update(verifier).digest('base64url');
}
