const express = require("express")
const cors = require("cors")

const app = express()
const PORT = 5000

app.use(cors())
app.use(express.json())

// Test simple
app.get("/api/status", (req, res) => {
  console.log("📋 Requête status reçue")
  res.json({
    message: "Serveur fonctionne!",
    timestamp: new Date().toISOString(),
    port: PORT,
  })
})

app.get("/api/tickets", (req, res) => {
  console.log("📊 Requête tickets reçue")
  res.json([])
})

app.listen(PORT, () => {
  console.log(`✅ Serveur minimal démarré sur le port ${PORT}`)
  console.log(`🌐 Test: http://localhost:${PORT}/api/status`)
})
