const express = require("express")
const multer = require("multer")
const path = require("path")

const router = express.Router()
const upload = multer({ dest: "uploads/" })

router.post("/api/upload", upload.single("excelFile"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("Aucun fichier reçu")
    }

    console.log("📁 Fichier reçu:", req.file)

    // Exemple de lecture avec XLSX
    const XLSX = require("xlsx")
    const workbook = XLSX.readFile(req.file.path)

    // tu peux faire tes traitements ici...
    res.json({ message: "Fichier uploadé et lu avec succès" })
  } catch (error) {
    console.error("❌ Erreur upload:", error)
    res.status(500).send("Erreur interne serveur")
  }
})

module.exports = router
