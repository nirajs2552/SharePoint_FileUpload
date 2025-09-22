
import 'dotenv/config'
import express from 'express'
import { makeMsalClient, oboExchange } from './obo.js'
import { downloadItem } from './graph.js'
import fetch from 'node-fetch'

const app = express()
app.use(express.json({ limit: '2mb' }))

const TENANT_ID = process.env.TENANT_ID
const BACKEND_CLIENT_ID = process.env.BACKEND_CLIENT_ID
const BACKEND_CLIENT_SECRET = process.env.BACKEND_CLIENT_SECRET
const BACKEND_REDIRECT_URI = process.env.BACKEND_REDIRECT_URI || 'http://localhost:4000/auth/redirect'
const INGEST_URL = process.env.INGEST_URL || 'http://localhost:5001/ingest'

const cca = makeMsalClient({
  tenantId: TENANT_ID,
  clientId: BACKEND_CLIENT_ID,
  clientSecret: BACKEND_CLIENT_SECRET,
  redirectUri: BACKEND_REDIRECT_URI
})

// Health
app.get('/health', (_, res) => res.json({ ok: true }))

/**
 * POST /api/ms/download-and-forward
 * Body: { items: [{ id, name, parentReference: { driveId } }], userToken?: string }
 * Frontend: send user access token via Authorization header OR in body.userToken.
 */
app.post('/api/ms/download-and-forward', async (req, res) => {
  try {
    const items = req.body.items || []
    const authHeader = req.headers['authorization']
    const bearer = authHeader?.startsWith('Bearer ') ? authHeader.slice(7) : (req.body.userToken || null)
    if (!bearer) return res.status(401).json({ error: 'Missing user token' })
    if (!items.length) return res.status(400).json({ error: 'No items' })

    // Exchange SPA token for Graph token (OBO). Request /.default to leverage app's Graph permissions.
    const { accessToken } = await oboExchange(cca, bearer, ['https://graph.microsoft.com/.default'])

    const results = []
    for (const it of items) {
      const driveId = it.parentReference?.driveId || it.driveId
      const itemId = it.id
      if (!driveId || !itemId) continue

      const { stream, name, size, meta } = await downloadItem(accessToken, driveId, itemId)

      // Forward to ingestion (streaming)
      const resp = await fetch(INGEST_URL + '?filename=' + encodeURIComponent(name), {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: stream
      })
      const ing = await resp.text()
      results.push({ name, size, ingested: resp.ok, ingestionResponse: ing, meta })
    }

    res.json({ ok: true, results })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: e.message })
  }
})

const port = parseInt(process.env.PORT || '4000', 10)
app.listen(port, () => console.log('Backend listening on', port))
