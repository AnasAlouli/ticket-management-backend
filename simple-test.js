  const express = require("express")
  const cors = require("cors")
  const multer = require("multer")
  const XLSX = require("xlsx")
  const fs = require("fs")
  const path = require("path")
  const moment = require("moment")

  const app = express()
  const PORT = 5000

  // Middleware
  app.use(cors())
  app.use(express.json())

  // Configuration des dossiers
  const EXCEL_DIR = path.join(__dirname, "excel")
  const UPLOADS_DIR = path.join(__dirname, "uploads")
  const TEMP_DIR = path.join(__dirname, "temp")

  // Créer les dossiers s'ils n'existent pas
  ;[EXCEL_DIR, UPLOADS_DIR, TEMP_DIR].forEach((dir) => {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
      console.log(`📂 Dossier créé: ${dir}`)
    }
  })

  // Configuration multer pour l'upload
  const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, UPLOADS_DIR)
    },
    filename: (req, file, cb) => {
      const timestamp = Date.now()
      const dateStr = moment(timestamp).format("YYYY-MM-DD_HH-mm-ss")
      const originalName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, "_")
      const newFilename = `${dateStr}_${originalName}`
      cb(null, newFilename)
    },
  })

  const upload = multer({ storage: storage })

  // Variables globales pour stockage temporaire
  let currentTicketsData = []
  let lastLoadedFile = null

  // Fonction pour lire un fichier Excel
  function readExcelFile(filePath) {
    try {
      if (!fs.existsSync(filePath)) return []

      const workbook = XLSX.readFile(filePath)
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      return XLSX.utils.sheet_to_json(worksheet)
    } catch (error) {
      console.error("Erreur lecture Excel:", error.message)
      return []
    }
  }

  // Fonction pour sauvegarder dans Excel, en écrasant l'ancien fichier
  function saveToExcel(data, originalFilename = "tickets.xlsx") {
    try {
      const timestamp = Date.now()
      const dateStr = moment(timestamp).format("YYYY-MM-DD_HH-mm-ss")
      const baseName = path.basename(originalFilename, path.extname(originalFilename))
      const ext = path.extname(originalFilename) || ".xlsx"
      const filename = `${dateStr}_${baseName}${ext}`
      const filePath = path.join(EXCEL_DIR, filename)

      // Supprimer tous les fichiers Excel existants dans EXCEL_DIR
      const existingFiles = fs.readdirSync(EXCEL_DIR).filter(f => f.match(/\.(xlsx|xls)$/i))
      existingFiles.forEach(f => {
        try {
          fs.unlinkSync(path.join(EXCEL_DIR, f))
        } catch {}
      })

      const worksheet = XLSX.utils.json_to_sheet(data)
      const workbook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(workbook, worksheet, "Tickets")
      XLSX.writeFile(workbook, filePath)

      if (fs.existsSync(filePath)) {
        return { success: true, filename: filename }
      }
      return { success: false, error: "Fichier non créé" }
    } catch (error) {
      console.error("Erreur sauvegarde Excel:", error.message)
      return { success: false, error: error.message }
    }
  }

  // Fonction pour charger le fichier Excel unique
  function loadMostRecentExcelFile() {
    try {
      const files = fs.readdirSync(EXCEL_DIR).filter(f => f.match(/\.(xlsx|xls)$/i))

      if (files.length === 0) return { data: [], file: null }

      // Comme il y a un seul fichier, on prend directement le premier
      const file = files[0]
      const filePath = path.join(EXCEL_DIR, file)
      const data = readExcelFile(filePath)
      return { data, file }
    } catch {
      return { data: [], file: null }
    }
  }


  // Charger les données principales
  function loadMainData() {
    const result = loadMostRecentExcelFile()
    currentTicketsData = result.data
    lastLoadedFile = result.file
    return result.data
  }

  // Routes API

  // GET tous tickets
// GET tous tickets avec option de recherche par Incident Number
app.get("/api/tickets", (req, res) => {
  try {
    const searchTerm = (req.query.search || "").toLowerCase().trim()
    const tickets = loadMainData()

    if (!searchTerm) {
      // Pas de recherche → tout renvoyer
      return res.json(tickets)
    }

    // Filtrage uniquement sur Incident Number
    const filtered = tickets.filter(ticket =>
      ticket["Incident Number"] &&
      ticket["Incident Number"].toString().toLowerCase().includes(searchTerm)
    )

    res.json(filtered)
  } catch (error) {
    console.error("Erreur chargement tickets:", error)
    res.status(500).json({ message: "Erreur chargement tickets" })
  }
})


  // POST création ticket sans validation
  app.post("/api/tickets", (req, res) => {
    try {
      const newTicket = req.body
      const tickets = loadMainData()
      // On n’empêche pas les doublons ni la validation
      tickets.push(newTicket)

      const saveResult = saveToExcel(tickets, lastLoadedFile || "tickets.xlsx")
      if (!saveResult.success) return res.status(500).json({ message: "Erreur sauvegarde" })

      lastLoadedFile = saveResult.filename
      currentTicketsData = tickets
      res.json({ message: "Ticket créé", ticket: newTicket })
    } catch {
      res.status(500).json({ message: "Erreur création ticket" })
    }
  })

  // PUT mise à jour ticket sans validation
  app.put("/api/tickets/:id", (req, res) => {
    try {
      const id = req.params.id
      const updatedData = req.body
      const tickets = loadMainData()
      const index = tickets.findIndex(t => t["Incident Number"] === id)
      if (index === -1) return res.status(404).json({ message: "Ticket non trouvé" })

      tickets[index] = { ...tickets[index], ...updatedData }
      const saveResult = saveToExcel(tickets, lastLoadedFile || "tickets.xlsx")
      if (!saveResult.success) return res.status(500).json({ message: "Erreur sauvegarde" })

      lastLoadedFile = saveResult.filename
      currentTicketsData = tickets
      res.json({ message: "Ticket mis à jour" })
    } catch {
      res.status(500).json({ message: "Erreur mise à jour ticket" })
    }
  })

  // DELETE ticket sans validation
  app.delete("/api/tickets/:id", (req, res) => {
    try {
      const id = req.params.id
      const tickets = loadMainData()
      const filtered = tickets.filter(t => t["Incident Number"] !== id)
      if (filtered.length === tickets.length) return res.status(404).json({ message: "Ticket non trouvé" })

      const saveResult = saveToExcel(filtered, lastLoadedFile || "tickets.xlsx")
      if (!saveResult.success) return res.status(500).json({ message: "Erreur sauvegarde" })

      lastLoadedFile = saveResult.filename
      currentTicketsData = filtered
      res.json({ message: "Ticket supprimé" })
    } catch {
      res.status(500).json({ message: "Erreur suppression ticket" })
    }
  })

  // POST upload fichier Excel sans validation
  app.post("/api/upload", upload.single("excelFile"), (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ message: "Aucun fichier fourni" })

      const uploadedFilePath = req.file.path
      if (!fs.existsSync(uploadedFilePath)) return res.status(500).json({ message: "Fichier uploadé introuvable" })

      const data = readExcelFile(uploadedFilePath)
      if (data.length === 0) {
        fs.unlinkSync(uploadedFilePath)
        return res.status(400).json({ message: "Fichier Excel vide ou illisible" })
      }

      // Remplace l'ancien fichier unique par celui-ci
      const saveResult = saveToExcel(data, req.file.originalname)
      if (!saveResult.success) return res.status(500).json({ message: "Erreur sauvegarde fichier" })

      lastLoadedFile = saveResult.filename
      currentTicketsData = data

      res.json({ message: "Fichier uploadé et sauvegardé", filename: saveResult.filename })
    } catch {
      res.status(500).json({ message: "Erreur upload fichier" })
    }
  })

  app.listen(PORT, () => {
    console.log(`🚀 Serveur lancé sur http://localhost:${PORT}`)
  })
