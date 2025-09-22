
import fetch from "node-fetch";

export async function getItemMetadata(graphToken, driveId, itemId) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`, {
    headers: { Authorization: `Bearer ${graphToken}` }
  });
  if (!res.ok) throw new Error(`Graph metadata error ${res.status}`);
  return res.json();
}

export async function downloadItem(graphToken, driveId, itemId) {
  // You can also stream from /content. Here we follow the downloadUrl for simplicity.
  const meta = await getItemMetadata(graphToken, driveId, itemId);
  const url = meta['@microsoft.graph.downloadUrl'];
  if (!url) throw new Error("No downloadUrl on item");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Download error ${res.status}`);
  return { stream: res.body, name: meta.name, size: meta.size || 0, meta };
}
