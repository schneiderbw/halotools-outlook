// PKCE helpers for the upstream Halo OAuth leg.
//
// MCP clients (Claude/Cursor) PKCE-protect their leg with us. We separately
// PKCE-protect our leg with Halo. Both legs use S256.

import { createHash, randomBytes } from "node:crypto";

function base64url(buf: Buffer): string {
  return buf
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

export function generateVerifier(): string {
  return base64url(randomBytes(32));
}

export function challengeFromVerifier(verifier: string): string {
  return base64url(createHash("sha256").update(verifier).digest());
}

/** Verify an inbound PKCE pair: the client sent code_challenge with the
 *  authorize call; later sends code_verifier with the token call. Both
 *  must agree under S256. We don't accept `plain`. */
export function verifyPkce(
  codeVerifier: string,
  codeChallenge: string,
  method: string,
): boolean {
  if (method !== "S256") return false;
  return challengeFromVerifier(codeVerifier) === codeChallenge;
}
