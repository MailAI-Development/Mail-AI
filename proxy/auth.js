/**
 * Verifies Microsoft identity-platform ID tokens sent by the web app
 * (Authorization: Bearer <id_token>), for work/school and personal Microsoft accounts.
 *
 * Checks: RS256 signature against Microsoft's published keys, audience = our app's client ID,
 * issuer = https://login.microsoftonline.com/<the token's own tenant>/v2.0, and expiry.
 */
import { createRemoteJWKSet, jwtVerify, decodeJwt } from "jose";

// Covers keys for all tenants, including personal Microsoft accounts. jose caches the key set.
const JWKS = createRemoteJWKSet(new URL("https://login.microsoftonline.com/common/discovery/v2.0/keys"));
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Returns {userId, tenantId, email, name} for a valid token, or null. */
export async function verifyMicrosoftToken(token, clientId) {
  if (!token || !clientId) return null;
  try {
    // The issuer depends on the user's tenant, so read it first, then require an exact match.
    const tid = decodeJwt(token).tid;
    if (typeof tid !== "string" || !GUID.test(tid)) return null;
    const { payload } = await jwtVerify(token, JWKS, {
      algorithms: ["RS256"],
      audience: clientId,
      issuer: `https://login.microsoftonline.com/${tid}/v2.0`,
      clockTolerance: 60,
    });
    return {
      userId: `${payload.tid}:${payload.oid || payload.sub}`,
      tenantId: payload.tid,
      email: payload.preferred_username || payload.email || null,
      name: payload.name || null,
    };
  } catch {
    return null;
  }
}
