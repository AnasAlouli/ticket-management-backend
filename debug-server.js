
  const express = require("express");
  const cors = require("cors");
  const multer = require("multer");
  const XLSX = require("xlsx");
  const fs = require("fs");
  const path = require("path");
  const moment = require("moment");

  const app = express();
  const PORT = 5000;

  // Middleware
  app.use(cors());
  app.use(express.json());

  // Configuration des dossiers
  const EXCEL_DIR = path.join(__dirname, "excel");
  const UPLOADS_DIR = path.join(__dirname, "uploads");
  const TEMP_DIR = path.join(__dirname, "temp");

  // Créer les dossiers s'ils n'existent pas
  [EXCEL_DIR, UPLOADS_DIR, TEMP_DIR].forEach((dir) => {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
      console.log(`📂 Dossier créé: ${dir}`);
    }
  });

  // Configuration multer pour l'upload
  const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, UPLOADS_DIR);
    },
    filename: (req, file, cb) => {
      const timestamp = Date.now();
      const dateStr = moment(timestamp).format("YYYY-MM-DD_HH-mm-ss");
      const originalName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, "_");
      const newFilename = `${dateStr}_${originalName}`;
      cb(null, newFilename);
    },
  });

  const upload = multer({ storage: storage });

  // Variables globales pour stockage temporaire
  let currentTicketsData = [];
  let lastLoadedFile = null;
  function excelDateToJSDate(serial) {
    if (typeof serial !== "number") return serial;
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);

    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    let total_seconds = Math.floor(86400 * fractional_day);

    const seconds = total_seconds % 60;
    total_seconds -= seconds;

    const hours = Math.floor(total_seconds / (60 * 60));
    const minutes = Math.floor(total_seconds / 60) % 60;

    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
  }
  // Fonction pour lire un fichier Excel
  function readExcelFile(filePath) {
    try {
      if (!fs.existsSync(filePath)) return [];

      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      let data = XLSX.utils.sheet_to_json(worksheet);

      // Conversion des dates Excel (exemple pour le champ "Date")
      data = data.map(row => {
        Object.keys(row).forEach(key => {
          // Si la valeur ressemble à une date Excel, convertis-la
          if (typeof row[key] === "number" && row[key] > 40000 && row[key] < 60000) {
            const jsDate = excelDateToJSDate(row[key]);
            // Format lisible
            row[key] = moment(jsDate).format("YYYY-MM-DD HH:mm:ss");
          }
        });
        return row;
      });

      return data;
    } catch (error) {
      console.error("Erreur lecture Excel:", error.message);
      return [];
    }
  }

  // Fonction pour sauvegarder dans Excel
  function saveToExcel(data, originalFilename = "tickets.xlsx") {
    try {
      const timestamp = Date.now();
      const dateStr = moment(timestamp).format("YYYY-MM-DD_HH-mm-ss");
      const baseName = path.basename(originalFilename, path.extname(originalFilename));
      const ext = path.extname(originalFilename) || ".xlsx";
      const filename = `${dateStr}_${baseName}${ext}`;
      const filePath = path.join(EXCEL_DIR, filename);

      // Supprimer tous les fichiers Excel existants dans EXCEL_DIR
      const existingFiles = fs.readdirSync(EXCEL_DIR).filter(f => f.match(/\.(xlsx|xls)$/i));
      existingFiles.forEach(f => {
        try {
          fs.unlinkSync(path.join(EXCEL_DIR, f));
        } catch {}
      });

      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Tickets");
      XLSX.writeFile(workbook, filePath);

      if (fs.existsSync(filePath)) {
        return { success: true, filename: filename };
      }
      return { success: false, error: "Fichier non créé" };
    } catch (error) {
      console.error("Erreur sauvegarde Excel:", error.message);
      return { success: false, error: error.message };
    }
  }

  // Fonction pour charger le fichier Excel unique
  function loadMostRecentExcelFile() {
    try {
      const files = fs.readdirSync(EXCEL_DIR).filter(f => f.match(/\.(xlsx|xls)$/i));

      if (files.length === 0) return { data: [], file: null };

      // Comme il y a un seul fichier, on prend directement le premier
      const file = files[0];
      const filePath = path.join(EXCEL_DIR, file);
      const data = readExcelFile(filePath);
      return { data, file };
    } catch {
      return { data: [], file: null };
    }
  }

  // Charger les données principales
  function loadMainData() {
    const result = loadMostRecentExcelFile();
    currentTicketsData = result.data;
    lastLoadedFile = result.file;
    return result.data;
  }

  // Fonction pour calculer les statistiques
  function calculateStats(tickets) {
    return {
      total: tickets.length,
      open: tickets.filter(t => t["Status"]?.toLowerCase() === "open").length,
      inProgress: tickets.filter(t => t["Status"]?.toLowerCase() === "in progress").length,
      resolved: tickets.filter(t => t["Status"]?.toLowerCase() === "closed").length,
    
    };
  }

  function getCountByField(tickets, fieldName) {
    const countMap = {};
    tickets.forEach(ticket => {
      const value = ticket[fieldName] || "Non spécifié";
      countMap[value] = (countMap[value] || 0) + 1;
    });
    return countMap;
  }

  // Routes API
  app.get("/api/next-incident-number", (req, res) => {
    try {
      const tickets = loadMainData();
      const existingNumbers = tickets
        .map(t => t["Incident Number"])
        .filter(Boolean);

      // Extraire les valeurs numériques
      const numericValues = [];
      existingNumbers.forEach(num => {
        const match = num.match(/^I(\d{4})-(\d{1,4})$/);
        if (!match) return;
        const prefix = match[1];
        const suffix = match[2].padStart(4, '0');
        const numericValue = parseInt(prefix + suffix, 10);
        if (!isNaN(numericValue)) numericValues.push(numericValue);
      });

      // Calculer le prochain numéro
      let nextNumber;
      if (numericValues.length > 0) {
        const maxValue = Math.max(...numericValues);
        const nextValue = maxValue + 1;
        const prefix = nextValue.toString().slice(0, 4);
        const suffix = nextValue.toString().slice(4).padStart(4, '0');
        nextNumber = `I${prefix}-${suffix}`;
      } else {
        // Fallback si aucun ticket
        const now = new Date();
        const yearPart = now.getFullYear().toString().slice(-2);
        const monthPart = (now.getMonth() + 1).toString().padStart(2, '0');
        const randomPart = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
        nextNumber = `I${yearPart}${monthPart}-${randomPart}`;
      }

      res.json({ nextIncidentNumber: nextNumber });
    } catch (error) {
      res.status(500).json({ message: "Erreur génération numéro", error: error.message });
    }
  });
  app.get("/api/tickets/all", (req, res) => {
  try {
    const tickets = loadMainData(); // charge tout
    res.json(tickets); // retourne la liste complète
  } catch (error) {
    console.error("Erreur chargement tickets:", error);
    res.status(500).json({ message: "Erreur chargement tickets" });
  }
});
  // GET tous tickets avec pagination et recherche
  app.get("/api/tickets", (req, res) => {
    try {

      const page = parseInt(req.query.page) || 1;
      const pageSize = parseInt(req.query.pageSize) || 50;
      
      const tickets = loadMainData();
      let filtered = tickets;

      /*if (searchTerm) {
        filtered = tickets.filter(ticket =>
          ticket["Incident Number"]?.toString().toLowerCase().includes(searchTerm) 
          
        );
      }*/
    

      const startIndex = (page - 1) * pageSize;
      const endIndex = startIndex + pageSize;
      const paginatedData = filtered.slice(startIndex, endIndex);

      res.json({
        tickets: paginatedData,
        total: filtered.length,
        page,
        totalPages: Math.ceil(filtered.length / pageSize),
        hasMore: endIndex < filtered.length
      });
    } catch (error) {
      console.error("Erreur chargement tickets:", error);
      res.status(500).json({ message: "Erreur chargement tickets" });
    }
  });

  // GET statistiques
  app.get("/api/stats", (req, res) => {
    try {
      const tickets = loadMainData();
      const stats = calculateStats(tickets);
      res.json(stats);
    } catch (error) {
      console.error("Erreur calcul stats:", error);
      res.status(500).json({ message: "Erreur calcul statistiques" });
    }
  });

  // GET ticket par ID
  app.get("/api/tickets/:id", (req, res) => {
    try {
      const id = req.params.id;
      const tickets = loadMainData();
      const ticket = tickets.find(t => t["Incident Number"] === id);
      
      if (!ticket) {
        return res.status(404).json({ message: "Ticket non trouvé" });
      }
      
      res.json(ticket);
    } catch (error) {
      console.error("Erreur recherche ticket:", error);
      res.status(500).json({ message: "Erreur recherche ticket" });
    }
  });

  // POST création ticket
  app.post("/api/tickets", (req, res) => {
    try {
      const newTicket = req.body;
      const tickets = loadMainData();
      
      // Vérifier si le ticket existe déjà
      const exists = tickets.some(t => t["Incident Number"] === newTicket["Incident Number"]);
      if (exists) {
        return res.status(400).json({ message: "Un ticket avec ce numéro existe déjà" });
      }
      
      tickets.push(newTicket);
      const saveResult = saveToExcel(tickets, lastLoadedFile || "tickets.xlsx");
      
      if (!saveResult.success) {
        return res.status(500).json({ message: "Erreur sauvegarde" });
      }
      
      lastLoadedFile = saveResult.filename;
      currentTicketsData = tickets;
      res.status(201).json({ message: "Ticket créé", ticket: newTicket });
    } catch (error) {
      console.error("Erreur création ticket:", error);
      res.status(500).json({ message: "Erreur création ticket" });
    }
  });

  // PUT mise à jour ticket
  app.put("/api/tickets/:id", (req, res) => {
    try {
      const id = req.params.id;
      const updatedData = req.body;
      const tickets = loadMainData();
      const index = tickets.findIndex(t => t["Incident Number"] === id);
      
      if (index === -1) {
        return res.status(404).json({ message: "Ticket non trouvé" });
      }
      
      tickets[index] = { ...tickets[index], ...updatedData };
      const saveResult = saveToExcel(tickets, lastLoadedFile || "tickets.xlsx");
      
      if (!saveResult.success) {
        return res.status(500).json({ message: "Erreur sauvegarde" });
      }
      
      lastLoadedFile = saveResult.filename;
      currentTicketsData = tickets;
      res.json({ message: "Ticket mis à jour", ticket: tickets[index] });
    } catch (error) {
      console.error("Erreur mise à jour ticket:", error);
      res.status(500).json({ message: "Erreur mise à jour ticket" });
    }
  });

  // DELETE ticket
  app.delete("/api/tickets/:id", (req, res) => {
    try {
      const id = req.params.id;
      const tickets = loadMainData();
      const filtered = tickets.filter(t => t["Incident Number"] !== id);
      
      if (filtered.length === tickets.length) {
        return res.status(404).json({ message: "Ticket non trouvé" });
      }
      
      const saveResult = saveToExcel(filtered, lastLoadedFile || "tickets.xlsx");
      
      if (!saveResult.success) {
        return res.status(500).json({ message: "Erreur sauvegarde" });
      }
      
      lastLoadedFile = saveResult.filename;
      currentTicketsData = filtered;
      res.json({ message: "Ticket supprimé" });
    } catch (error) {
      console.error("Erreur suppression ticket:", error);
      res.status(500).json({ message: "Erreur suppression ticket" });
    }
  });

  // POST upload fichier Excel
  app.post("/api/upload", upload.single("excelFile"), (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ message: "Aucun fichier fourni" });
      }

      const uploadedFilePath = req.file.path;
      if (!fs.existsSync(uploadedFilePath)) {
        return res.status(500).json({ message: "Fichier uploadé introuvable" });
      }

      const data = readExcelFile(uploadedFilePath);
      if (data.length === 0) {
        fs.unlinkSync(uploadedFilePath);
        return res.status(400).json({ message: "Fichier Excel vide ou illisible" });
      }

      // Supprimer l'ancien fichier et sauvegarder le nouveau
      const saveResult = saveToExcel(data, req.file.originalname);
      if (!saveResult.success) {
        return res.status(500).json({ message: "Erreur sauvegarde fichier" });
      }

      // Nettoyer le fichier uploadé
      try {
        fs.unlinkSync(uploadedFilePath);
      } catch {}

      lastLoadedFile = saveResult.filename;
      currentTicketsData = data;

      res.json({ 
        message: "Fichier uploadé et sauvegardé", 
        filename: saveResult.filename,
        stats: calculateStats(data)
      });
    } catch (error) {
      console.error("Erreur upload fichier:", error);
      res.status(500).json({ message: "Erreur upload fichier" });
    }
  });

  // GET liste des fichiers disponibles
  app.get("/api/files", (req, res) => {
    try {
      const files = fs.readdirSync(EXCEL_DIR)
        .filter(f => f.match(/\.(xlsx|xls)$/i))
        .map(f => ({
          filename: f,
          path: path.join(EXCEL_DIR, f),
          size: fs.statSync(path.join(EXCEL_DIR, f)).size,
          lastModified: fs.statSync(path.join(EXCEL_DIR, f)).mtime
        }));

      res.json(files);
    } catch (error) {
      console.error("Erreur liste fichiers:", error);
      res.status(500).json({ message: "Erreur liste fichiers" });
    }
  });

  // GET exporter les données actuelles
  app.get("/api/export", (req, res) => {
    try {
      const tickets = loadMainData();
      if (tickets.length === 0) {
        return res.status(404).json({ message: "Aucune donnée à exporter" });
      }

      // Ajout du nom du fichier source dans le nom d'export
      const originalName = lastLoadedFile ? path.basename(lastLoadedFile, path.extname(lastLoadedFile)) : "tickets";
      const ext = ".xlsx";
      const dateStr = moment().format("YYYY-MM-DD_HH-mm-ss");
      const filename = `export_${dateStr}_${originalName}${ext}`;
      const filePath = path.join(TEMP_DIR, filename);

      const worksheet = XLSX.utils.json_to_sheet(tickets);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Tickets");
      XLSX.writeFile(workbook, filePath);

      res.download(filePath, filename, (err) => {
        if (err) {
          console.error("Erreur envoi fichier:", err);
        }
        // Nettoyer après envoi
        try {
          fs.unlinkSync(filePath);
        } catch {}
      });
    } catch (error) {
      console.error("Erreur export:", error);
      res.status(500).json({ message: "Erreur export données" });
    }
  });

  // Démarrer le serveur
  app.listen(PORT, () => {
    console.log(`🚀 Serveur lancé sur http://localhost:${PORT}`);
    // Charger les données au démarrage
    loadMainData();
  });
