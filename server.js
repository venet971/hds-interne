const express = require('express');
const session = require('express-session');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const { Document, Packer, Paragraph, TextRun, ImageRun } = require('docx');

const app = express();
const PORT = process.env.PORT || 3000;

/* ================= CONFIG ================= */

app.use(session({
  secret: 'hds-secret',
  resave: false,
  saveUninitialized: false
}));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

/* ================= UPLOAD ================= */

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/'),
  filename: (req, file, cb) => {
    const name = Date.now() + '-' + file.originalname.replace(/\s+/g, '_');
    cb(null, name);
  }
});

const upload = multer({ storage });

/* ================= DATA ================= */

function droitsParDefaut() {
  return {
    chantiers: true,
    depenses: true,
    feuillesHeures: true,
    interventions: true,
    photos: true,
    reglages: true
  };
}

function lireDonnees() {
  try {
    const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

    return {
      chantiers: data.chantiers || [],
      depenses: data.depenses || [],
      utilisateurs: data.utilisateurs || [],
      photos: data.photos || [],
      feuillesHeures: data.feuillesHeures || [],
      interventions: data.interventions || [],
      archives: data.archives || [],
      reglages: {
        pourcentageFraisGeneraux: data.reglages?.pourcentageFraisGeneraux ?? 8,
        prixHeureCadre: data.reglages?.prixHeureCadre ?? 75,
        prixHeureTechnicien: data.reglages?.prixHeureTechnicien ?? 55
      }
    };

  } catch {
    return {
      chantiers: [],
      depenses: [],
      utilisateurs: [],
      photos: [],
      feuillesHeures: [],
      interventions: [],
      archives: [],
      reglages: {
        pourcentageFraisGeneraux: 8,
        prixHeureCadre: 75,
        prixHeureTechnicien: 55
      }
    };
  }
}

function sauvegarder(data) {
  fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
}

/* ================= AUTH ================= */

function verifierConnexion(req, res, next) {
  if (!req.session.user) return res.redirect('/');
  next();
}

/* ================= ROUTES ================= */

app.get('/', (req, res) => {
  if (!req.session.user) {
    res.send(pageConnexion());
    return;
  }

  const data = lireDonnees();
  const userActuel = data.utilisateurs.find(u => u.nom === req.session.user.nom) || req.session.user;
  req.session.user = userActuel;

  const message = req.query.message || '';
  res.send(pageDashboard(data, userActuel, message));
});

/* ================= LOGIN ================= */

app.post('/login', (req, res) => {
  const data = lireDonnees();

  const user = data.utilisateurs.find(
    u => u.nom === req.body.nom && u.motdepasse === req.body.motdepasse
  );

  if (!user) return res.send('Erreur login');

  req.session.user = user;
  res.redirect('/');
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => res.redirect('/'));
});

/* ================= CHANTIER ================= */

app.post('/chantier', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const numero = req.body.numero;

  if (!numero) return res.status(400).send('Numéro obligatoire');

  data.chantiers.push({
    numero,
    nom: req.body.nom || '',
    date: req.body.date || '',
    maitreOuvrage: req.body.maitreOuvrage || '',
    maitreOeuvre: req.body.maitreOeuvre || '',
    lieu: req.body.lieu || '',
    totalMarcheHT: Number(req.body.totalMarcheHT) || 0
  });

  sauvegarder(data);
  res.send('OK');
});

/* ================= ARCHIVAGE ================= */

app.post('/supprimer-chantier', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const numero = req.body.numero;
  const archiver = req.body.archiver === true;

  const chantier = data.chantiers.find(c => c.numero === numero);
  if (!chantier) return res.send('Introuvable');

  if (archiver) {
    data.archives.push({
      id: Date.now(),
      dateArchivage: new Date().toISOString(),
      numero,
      chantier,
      depenses: data.depenses.filter(d => d.numero === numero),
      feuillesHeures: data.feuillesHeures.filter(f => f.numero === numero),
      interventions: data.interventions.filter(i => i.numero === numero),
      photos: data.photos.filter(p => p.chantier === numero)
    });
  }

  data.chantiers = data.chantiers.filter(c => c.numero !== numero);
  data.depenses = data.depenses.filter(d => d.numero !== numero);
  data.feuillesHeures = data.feuillesHeures.filter(f => f.numero !== numero);
  data.interventions = data.interventions.filter(i => i.numero !== numero);
  data.photos = data.photos.filter(p => p.chantier !== numero);

  sauvegarder(data);
  res.send('OK');
});

/* ================= RESTAURATION ================= */

app.post('/restaurer-archive', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const id = req.body.id;

  const archive = data.archives.find(a => String(a.id) === String(id));
  if (!archive) return res.send('Archive introuvable');

  data.chantiers.push(archive.chantier);
  data.depenses.push(...archive.depenses);
  data.feuillesHeures.push(...archive.feuillesHeures);
  data.interventions.push(...archive.interventions);
  data.photos.push(...archive.photos);

  data.archives = data.archives.filter(a => a.id != id);

  sauvegarder(data);
  res.send('OK');
});

/* ================= PAGE ARCHIVE ================= */

app.get('/archive/:id', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const archive = data.archives.find(a => String(a.id) === String(req.params.id));
  if (!archive) return res.send('Archive introuvable');

  const totalMarche = Number(archive.chantier?.totalMarcheHT) || 0;

  const totalDepenses = (archive.depenses || []).reduce((sum, d) => {
    return sum + (Number(d.quantite) || 0) * (Number(d.montant) || 0);
  }, 0);

  const reste = totalMarche - totalDepenses;

  res.send(`
    <h1>Archive ${archive.numero}</h1>
    <p>Marché : ${totalMarche} €</p>
    <p>Dépenses : ${totalDepenses} €</p>
    <p>Reste : ${reste} €</p>
  `);
});

/* ================= PHOTO ================= */

app.post('/photo', verifierConnexion, upload.single('photo'), (req, res) => {
  const data = lireDonnees();

  data.photos.push({
    chantier: req.body.chantier,
    description: req.body.description,
    fichier: '/uploads/' + req.file.filename
  });

  sauvegarder(data);
  res.redirect('/');
});

/* ================= START ================= */

app.listen(PORT, () => {
  console.log('HDS lancé sur port ' + PORT);
});