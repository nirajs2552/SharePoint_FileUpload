export async function gfetch(token, url, init = {}) {
  const res = await fetch(url, {
    ...init,
    headers: { ...(init.headers || {}), Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${url}`);
  return res.json();
}

/** Normalize any SharePoint URL to { host, kind, siteName, siteRootUrl }
 *  Examples:
 *   https://contoso.sharepoint.com/sites/TimeSheet/Shared%20Documents/foo.pdf
 *   -> { host: 'contoso.sharepoint.com', kind: 'sites', siteName: 'TimeSheet',
 *        siteRootUrl: 'https://contoso.sharepoint.com/sites/TimeSheet' }
 *
 *   https://contoso.sharepoint.com/teams/HR/SiteAssets/logo.png
 *   -> { host: 'contoso.sharepoint.com', kind: 'teams', siteName: 'HR',
 *        siteRootUrl: 'https://contoso.sharepoint.com/teams/HR' }
 */
export function normalizeSiteFromUrl(raw) {
  try {
    const u = new URL(raw);
    const m =
      u.pathname.match(/\/(sites|teams)\/([^\/]+)/i); // grab first segment after /sites or /teams
    if (!m) return null;
    const kind = m[1];
    const siteName = decodeURIComponent(m[2]);
    const host = u.host; // e.g., contoso.sharepoint.com
    const siteRootUrl = `${u.protocol}//${host}/${kind}/${siteName}`;
    return { host, kind, siteName, siteRootUrl };
  } catch {
    return null;
  }
}

/** Resolve a site ID using Graph's "by path" endpoint */
export async function resolveSiteIdByPath(token, host, kind, siteName) {
  // GET /v1.0/sites/{host}:/sites/{siteName}
  const pathSegment = kind.toLowerCase() === "teams" ? "teams" : "sites";
  const url = `https://graph.microsoft.com/v1.0/sites/${host}:/${pathSegment}/${encodeURIComponent(siteName)}`;
  const data = await gfetch(token, url);
  return data.id; // hostname,siteCollectionId,siteId
}

/** Resolve site id from ANY SharePoint URL (file/page/library/site) */
export async function resolveSiteIdFromAnyUrl(token, anyUrl) {
  const norm = normalizeSiteFromUrl(anyUrl);
  if (!norm) throw new Error("Could not parse site from url");
  return resolveSiteIdByPath(token, norm.host, norm.kind, norm.siteName);
}

/** Candidate sites:
 *  - me/insights/used (fast, relevant)
 *  - me/followedSites (fallback)
 *  Return a list of unique *site root URLs*.
 */
export async function listCandidateSites(token) {
  const candidates = new Set();

  // 1) Recent/used
  try {
    const used = await gfetch(token, "https://graph.microsoft.com/v1.0/me/insights/used?$top=50");
    for (const v of used.value || []) {
      const url = v.resourceReference?.webUrl || v.resourceVisualization?.previewImageUrl;
      if (!url) continue;
      const norm = normalizeSiteFromUrl(url);
      if (norm?.siteRootUrl) candidates.add(norm.siteRootUrl);
    }
  } catch {/* ignore */}

  // 2) Followed
  try {
    const f = await gfetch(token, "https://graph.microsoft.com/v1.0/me/followedSites?$top=50");
    for (const s of f.value || []) {
      const norm = normalizeSiteFromUrl(s.webUrl);
      if (norm?.siteRootUrl) candidates.add(norm.siteRootUrl);
    }
  } catch {/* ignore */}

  return Array.from(candidates);
}

/** Drives (document libraries) for a site id */
export async function listSiteDrives(token, siteId) {
  const data = await gfetch(token, `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`);
  return data.value || [];
}

/** List children at drive root */
export async function listDriveRootChildren(token, driveId) {
  const data = await gfetch(token, `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=200`);
  return data.value || [];
}

/** List children under a folder item */
export async function listItemChildren(token, driveId, itemId) {
  const data = await gfetch(token, `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$top=200`);
  return data.value || [];
}
