import React, { useEffect, useState } from "react";
import {
  listCandidateSites,
  resolveSiteIdFromAnyUrl,
  listSiteDrives,
  listDriveRootChildren,
  listItemChildren,
} from "./msgraph";

/**
 * Props:
 *  - getToken(): Promise<string>  // returns MSAL access token
 *  - onPick(items: Array<{id, name, parentReference: { driveId }, file? }>)
 */
export default function Explorer({ getToken, onPick }) {
  const [loading, setLoading] = useState(false);
  const [token, setToken] = useState(null);

  // Sites
  const [sites, setSites] = useState([]);          // array of site root URLs
  const [selectedSiteUrl, setSelectedSiteUrl] = useState("");
  const [siteId, setSiteId] = useState("");        // Graph site id

  // Libraries (drives)
  const [drives, setDrives] = useState([]);
  const [driveId, setDriveId] = useState("");

  // Navigation (folders/files)
  const [path, setPath] = useState([]);            // breadcrumb of items [{id,name}]
  const [children, setChildren] = useState([]);    // items at current folder

  // Selection
  const [picked, setPicked] = useState([]);        // selected files (not folders)

  // Bootstrap: token + candidate sites
  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const t = await getToken();
        setToken(t);
        const s = await listCandidateSites(t);
        setSites(s);
      } catch (e) {
        console.error(e);
      } finally {
        setLoading(false);
      }
    })();
  }, [getToken]);

  async function selectSite(webUrl) {
    try {
      setSelectedSiteUrl(webUrl);
      setLoading(true);
      const id = await resolveSiteIdFromAnyUrl(token, webUrl);
      setSiteId(id);
      const d = await listSiteDrives(token, id);
      setDrives(d);
      // reset lower levels
      setDriveId("");
      setPath([]);
      setChildren([]);
      setPicked([]);
    } catch (e) {
      alert("Failed to resolve site: " + e.message);
    } finally {
      setLoading(false);
    }
  }

  async function selectDrive(id) {
    try {
      setDriveId(id);
      setLoading(true);
      const rootChildren = await listDriveRootChildren(token, id);
      setPath([]); // at root
      setChildren(rootChildren);
    } catch (e) {
      alert("Failed to list drive root: " + e.message);
    } finally {
      setLoading(false);
    }
  }

  async function enterFolder(item) {
    try {
      setLoading(true);
      const next = await listItemChildren(token, driveId, item.id);
      setPath((prev) => [...prev, { id: item.id, name: item.name }]);
      setChildren(next);
    } catch (e) {
      alert("Failed to open folder: " + e.message);
    } finally {
      setLoading(false);
    }
  }

  async function goBreadcrumb(index) {
    try {
      setLoading(true);
      if (index < 0) {
        const root = await listDriveRootChildren(token, driveId);
        setPath([]);
        setChildren(root);
        return;
      }
      const target = path[index];
      const kids = await listItemChildren(token, driveId, target.id);
      setPath(path.slice(0, index + 1));
      setChildren(kids);
    } catch (e) {
      alert("Failed to navigate: " + e.message);
    } finally {
      setLoading(false);
    }
  }

  function togglePick(item) {
    if (item.folder) return; // only pick files
    const key = `${item.parentReference?.driveId || driveId}:${item.id}`;
    setPicked((prev) => {
      const exists = prev.find(
        (p) =>
          (p.parentReference?.driveId || "") + ":" + p.id === key
      );
      if (exists)
        return prev.filter(
          (p) =>
            (p.parentReference?.driveId || "") + ":" + p.id !== key
        );
      return [
        ...prev,
        {
          id: item.id,
          name: item.name,
          parentReference: {
            driveId: item.parentReference?.driveId || driveId,
          },
          file: item.file,
        },
      ];
    });
  }

  return (
    <div style={{ border: "1px solid #ddd", borderRadius: 12, padding: 16 }}>
      <h3 style={{ marginTop: 0 }}>
        Microsoft Explorer (Sites → Libraries → Folders → Files)
      </h3>
      {loading && <div>Loading…</div>}

      {/* Sites */}
      <div style={{ marginTop: 8 }}>
        <div style={{ fontWeight: 600, marginBottom: 4 }}>
          Sites you’ve used/followed
        </div>
        {!sites.length && (
          <div style={{ color: "#666" }}>
            No recent/followed sites found. Paste a SharePoint site URL below
            and press Enter.
          </div>
        )}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {sites.map((u) => (
            <button
              key={u}
              onClick={() => selectSite(u)}
              title={u}
              style={{
                padding: "6px 10px",
                borderRadius: 8,
                border: "1px solid #ccc",
                background: selectedSiteUrl === u ? "#eef6ff" : "white",
                maxWidth: 380,
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
            >
              {u.replace(/^https?:\/\//, "")}
            </button>
          ))}
        </div>
        <input
          placeholder="Or paste a SharePoint site URL and press Enter"
          style={{ marginTop: 8, width: "100%", padding: 8 }}
          onKeyDown={(e) => {
            if (e.key === "Enter" && e.currentTarget.value)
              selectSite(e.currentTarget.value.trim());
          }}
        />
      </div>

      {/* Drives (Document Libraries) */}
      {!!siteId && (
        <div style={{ marginTop: 16 }}>
          <div style={{ fontWeight: 600, marginBottom: 4 }}>
            Document Libraries
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {drives.map((d) => (
              <button
                key={d.id}
                onClick={() => selectDrive(d.id)}
                style={{
                  padding: "6px 10px",
                  borderRadius: 8,
                  border: "1px solid #ccc",
                  background: driveId === d.id ? "#eef6ff" : "white",
                }}
                title={d.webUrl}
              >
                {d.name}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Breadcrumb */}
      {!!driveId && (
        <div style={{ marginTop: 16 }}>
          <div style={{ fontWeight: 600, marginBottom: 4 }}>Path</div>
          <div
            style={{
              display: "flex",
              gap: 6,
              alignItems: "center",
              flexWrap: "wrap",
            }}
          >
            <a
              href="#/"
              onClick={(e) => {
                e.preventDefault();
                goBreadcrumb(-1);
              }}
            >
              Root
            </a>
            {path.map((p, i) => (
              <React.Fragment key={p.id}>
                <span>›</span>
                <a
                  href="#/"
                  onClick={(e) => {
                    e.preventDefault();
                    goBreadcrumb(i);
                  }}
                >
                  {p.name}
                </a>
              </React.Fragment>
            ))}
          </div>
        </div>
      )}

      {/* Children list */}
      {!!driveId && (
        <div
          style={{
            marginTop: 8,
            maxHeight: 360,
            overflow: "auto",
            border: "1px solid #eee",
            borderRadius: 8,
          }}
        >
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={{ textAlign: "left", padding: 8 }}>Name</th>
                <th style={{ textAlign: "left", padding: 8 }}>Type</th>
                <th style={{ textAlign: "left", padding: 8 }}>Pick</th>
              </tr>
            </thead>
            <tbody>
              {children.map((item) => (
                <tr key={item.id} style={{ borderTop: "1px solid #f0f0f0" }}>
                  <td style={{ padding: 8 }}>
                    {item.folder ? (
                      <a
                        href="#/"
                        onClick={(e) => {
                          e.preventDefault();
                          enterFolder(item);
                        }}
                      >
                        📁 {item.name}
                      </a>
                    ) : (
                      <span>📄 {item.name}</span>
                    )}
                  </td>
                  <td style={{ padding: 8 }}>
                    {item.folder ? "Folder" : item.file?.mimeType || "File"}
                  </td>
                  <td style={{ padding: 8 }}>
                    {!item.folder && (
                      <button onClick={() => togglePick(item)}>
                        {picked.find((p) => p.id === item.id)
                          ? "Unpick"
                          : "Pick"}
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Actions */}
      <div
        style={{
          marginTop: 12,
          display: "flex",
          gap: 8,
          alignItems: "center",
        }}
      >
        <button onClick={() => onPick(picked)} disabled={!picked.length}>
          Use {picked.length} selected
        </button>
        <span style={{ color: "#666" }}>
          Files only (folders are for navigation)
        </span>
      </div>
    </div>
  );
}
export async function getItemDownloadInfo(token, driveId, itemId) {
  // Ask Graph for the item and the pre-signed download URL
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}?$select=name,@microsoft.graph.downloadUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Graph ${res.status} for ${driveId}/${itemId}`);
  const data = await res.json();
  const url = data['@microsoft.graph.downloadUrl'];
  if (!url) throw new Error('No downloadUrl on item');
  return { name: data.name, downloadUrl: url };
}

