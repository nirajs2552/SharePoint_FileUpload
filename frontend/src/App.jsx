import React, { useEffect, useState } from 'react'
import { getAccessToken, signIn, signOut, getAccount } from './msal'
import Explorer from './Explorer'

const backendUrl = import.meta.env.VITE_BACKEND_URL

export default function App() {
  const [items, setItems] = useState([])
  const [activeTab, setActiveTab] = useState('microsoft') // 'direct' | 's3' | 'microsoft' | 'signout'
  const [account, setAccount] = useState(null)

  useEffect(() => {
    (async () => {
      const acc = await getAccount()
      setAccount(acc)
    })()
  }, [])

  async function handlePick(picked) {
    setItems(picked)
  }

  async function ensureSignedIn() {
    const acc = await signIn()
    setAccount(acc)
  }

  async function uploadSelected() {
    if (!items.length) return alert("No items selected")
    const token = await getAccessToken()
    const res = await fetch(`${backendUrl}/api/ms/download-and-forward`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({ items })
    })
    const data = await res.json()
    alert("Ingested: " + JSON.stringify(data, null, 2))
  }

  async function handleSignOut() {
    await signOut()
    setItems([])
    setAccount(null)
    setActiveTab('microsoft')
  }

  return (
    <div style={{ fontFamily: 'sans-serif', padding: 24 }}>
      <h1>AgentFleet — Upload</h1>

      {/* Tabs */}
      <div style={{ display: 'flex', gap: 12, marginBottom: 16 }}>
        <Tab label="Direct Upload" active={activeTab === 'direct'} onClick={() => setActiveTab('direct')} />
        <Tab label="S3 Upload" active={activeTab === 's3'} onClick={() => setActiveTab('s3')} />
        <Tab label="Microsoft" active={activeTab === 'microsoft'} onClick={() => setActiveTab('microsoft')} />
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}>
          {account ? (
            <>
              <span style={{ color: '#555' }}>{account.name || account.username}</span>
              <Tab label="Sign out" onClick={handleSignOut} />
            </>
          ) : (
            <Tab label="Sign in" onClick={ensureSignedIn} />
          )}
        </div>
      </div>

      {/* Content area */}
      {activeTab === 'microsoft' && (
        <>
          {!account && (
            <div style={{ marginBottom: 12 }}>
              <button onClick={ensureSignedIn}>Sign in with Microsoft</button>
            </div>
          )}

          <p>Browse your SharePoint/OneDrive (Sites → Libraries → Folders → Files), then upload.</p>
          <Explorer getToken={getAccessToken} onPick={handlePick} />

          <div style={{ marginTop: 12 }}>
            <button onClick={uploadSelected} disabled={!items.length}>Upload Selected</button>
          </div>

          <pre style={{ marginTop: 16, background: '#f6f6f6', padding: 12 }}>
            Selected Items: {JSON.stringify(items, null, 2)}
          </pre>
        </>
      )}

      {activeTab === 'direct' && (
        <div style={{ padding: 12, border: '1px solid #eee', borderRadius: 8 }}>
          <em>Direct upload UI goes here…</em>
        </div>
      )}

      {activeTab === 's3' && (
        <div style={{ padding: 12, border: '1px solid #eee', borderRadius: 8 }}>
          <em>S3 upload UI goes here…</em>
        </div>
      )}
    </div>
  )
}

function Tab({ label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: '8px 12px',
        borderRadius: 8,
        border: '1px solid ' + (active ? '#3b82f6' : '#ddd'),
        background: active ? '#eff6ff' : '#fff',
        cursor: 'pointer'
      }}
    >
      {label}
    </button>
  )
}
