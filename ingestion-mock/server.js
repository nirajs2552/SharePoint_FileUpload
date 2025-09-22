
import express from 'express'
const app = express()

app.post('/ingest', (req, res) => {
  let bytes = 0
  req.on('data', chunk => bytes += chunk.length)
  req.on('end', () => res.send(`received ${bytes} bytes`))
})

app.listen(5001, () => console.log('Ingestion mock on 5001'))
