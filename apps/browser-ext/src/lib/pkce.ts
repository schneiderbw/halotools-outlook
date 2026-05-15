// PKCE (RFC 7636) helpers. Pure browser, no deps.
//
// Code verifier: 43-128 chars from the unreserved set [A-Z a-z 0-9 - . _ ~].
// Code challenge: BASE64URL(SHA256(verifier)) — S256 method.
//
// COPIED verbatim from ../../../src/lib/pkce.ts. The browser-extension and
// Outlook variants are identical — `crypto.subtle` is available in both
// service workers and DOM contexts, so this file needs no adaptation.

/** Generate a cryptographically random PKCE verifier (96 chars). */
export function generateVerifier(): string {
  const bytes = new Uint8Array(72); // 72 bytes → 96 base64url chars
  crypto.getRandomValues(bytes);
  return base64urlEncode(bytes);
}

/** Compute the SHA-256 code_challenge for a given verifier. */
export async function challengeFromVerifier(verifier: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  return base64urlEncode(new Uint8Array(digest));
}

/** Random URL-safe state for CSRF protection on the OAuth dance. */
export function generateState(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return base64urlEncode(bytes);
}

function base64urlEncode(bytes: Uint8Array): string {
  let str = "";
  for (const b of bytes) str += String.fromCharCode(b);
  return btoa(str).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
