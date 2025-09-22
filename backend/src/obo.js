
import { ConfidentialClientApplication } from "@azure/msal-node";

const oboCache = new Map();

export function makeMsalClient({ tenantId, clientId, clientSecret, redirectUri }) {
  const config = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientSecret
    },
    system: { loggerOptions: { loggerCallback: () => {} } }
  };
  return new ConfidentialClientApplication(config);
}

/**
 * Exchanges a user's SPA access token for a Graph access token (OBO).
 * @param {ConfidentialClientApplication} cca
 * @param {string} userToken - the SPA access token (bearer) from MSAL.js
 * @param {string[]} scopes - Graph scopes to request, e.g., ["https://graph.microsoft.com/.default"]
 */
export async function oboExchange(cca, userToken, scopes) {
  const cacheKey = `${scopes.join(' ')}:${userToken.slice(-24)}`;
  if (oboCache.has(cacheKey)) {
    const cached = oboCache.get(cacheKey);
    if (cached.expiresOn > Date.now() + 60_000) return cached;
  }
  const result = await cca.acquireTokenOnBehalfOf({
    oboAssertion: userToken,
    scopes
  });
  const token = {
    accessToken: result.accessToken,
    expiresOn: result.expiresOn.getTime()
  };
  oboCache.set(cacheKey, token);
  return token;
}
