const session = require('express-session');
const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const { Document, Packer, Paragraph, TextRun, ImageRun } = require('docx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(session({
  secret: 'hds-secret',
  resave: false,
  saveUninitialized: false
}));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    const uniqueName = Date.now() + '-' + file.originalname.replace(/\s+/g, '_');
    cb(null, uniqueName);
  }
});

const upload = multer({ storage });

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

    if (!data.chantiers) data.chantiers = [];
    if (!data.depenses) data.depenses = [];
    if (!data.utilisateurs) data.utilisateurs = [];
    if (!data.photos) data.photos = [];
    if (!data.feuillesHeures) data.feuillesHeures = [];
    if (!data.interventions) data.interventions = [];
    if (!data.archives) data.archives = [];

    data.utilisateurs = data.utilisateurs.map(u => {
      if (!u.nom) u.nom = '';
      if (!u.motdepasse) u.motdepasse = '';
      if (!u.role) u.role = '';
      if (!u.droits) u.droits = droitsParDefaut();

      if (typeof u.droits.chantiers === 'undefined') u.droits.chantiers = true;
      if (typeof u.droits.depenses === 'undefined') u.droits.depenses = true;
      if (typeof u.droits.feuillesHeures === 'undefined') u.droits.feuillesHeures = true;
      if (typeof u.droits.interventions === 'undefined') u.droits.interventions = true;
      if (typeof u.droits.photos === 'undefined') u.droits.photos = true;
      if (typeof u.droits.reglages === 'undefined') u.droits.reglages = true;

      return u;
    });

    if (!data.reglages) {
      data.reglages = {
        pourcentageFraisGeneraux: 8,
        prixHeureCadre: 75,
        prixHeureTechnicien: 55
      };
    }

    if (typeof data.reglages.pourcentageFraisGeneraux === 'undefined') {
      data.reglages.pourcentageFraisGeneraux = 8;
    }

    if (typeof data.reglages.prixHeureCadre === 'undefined') {
      data.reglages.prixHeureCadre = 75;
    }

    if (typeof data.reglages.prixHeureTechnicien === 'undefined') {
      data.reglages.prixHeureTechnicien = 55;
    }

    return data;
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

function sauvegarder(donnees) {
  fs.writeFileSync('data.json', JSON.stringify(donnees, null, 2));
  fs.writeFileSync('backup.json', JSON.stringify(donnees, null, 2));
}

function pageConnexion(message = '') {
  return `
    <html>
    <head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Connexion HDS Interne</title>
      <style>
        body { font-family: Arial; background:#f3f4f6; margin:0; padding:0; }
        .header { text-align:center; background:#1e293b; padding:15px; }
        .header img { height:80px; }
        .header .titre { color:white; font-size:20px; margin-top:8px; }
        .box {
          max-width: 420px;
          margin: 60px auto;
          background: white;
          padding: 25px;
          border-radius: 12px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        input {
          width: 100%;
          padding: 10px;
          margin: 8px 0;
          box-sizing: border-box;
          border:1px solid #ccc;
          border-radius:6px;
        }
        button {
          width: 100%;
          padding: 10px;
          background:#2563eb;
          color:white;
          border:none;
          border-radius:6px;
          cursor:pointer;
        }
        button:hover { opacity:0.9; }
        .msg { color:red; margin-bottom:10px; }
      </style>
    </head>
    <body>
      <div class="menu-mobile">
  <button onclick="showBloc('chantiers')">Chantiers</button>
  <button onclick="showBloc('depenses')">Dépenses</button>
  <button onclick="showBloc('interventions')">Interventions</button>
  <button onclick="showBloc('photos')">Photos</button>
  <button onclick="showBloc('reglages')">Réglages</button>
</div>
      <div class="header">
        <img src="/uploads/logo.jpg" alt="Logo Hydro Dom Solutions">
        <div class="titre">HDS Interne</div>
      </div>

      <div class="box">
        <h2>Connexion</h2>
        ${message ? `<div class="msg">${message}</div>` : ''}
        <form method="POST" action="/login">
          <input type="text" name="nom" placeholder="Nom" required>
          <input type="password" name="motdepasse" placeholder="Mot de passe" required>
          <button type="submit">Se connecter</button>
        </form>
      </div>
    </body>
    </html>
  `;
}

function pageDashboard(data, user, message = '') {
  const droits = user.droits || droitsParDefaut();
  const pourcentageFraisGeneraux = Number(data.reglages?.pourcentageFraisGeneraux) || 8;
  const prixHeureCadre = Number(data.reglages?.prixHeureCadre) || 75;
  const prixHeureTechnicien = Number(data.reglages?.prixHeureTechnicien) || 55;

  const optionsChantiers = data.chantiers.map(c =>
    `<option value="${c.numero || ''}" data-nom="${c.nom || ''}">
      ${c.numero || ''} - ${c.nom || ''}
    </option>`
  ).join('');

  let totalInitial = 0;
  let totalEncours = 0;
  let totalReste = 0;

  const htmlDepenses = data.depenses.map(d => {
    const quantite = Number(d.quantite) || 0;
    const montant = Number(d.montant) || 0;
    const totalLigne = quantite * montant;
    const estFraisGeneraux = d.type === 'Frais généraux' || d.verrouillee === true;
    const estMainDoeuvreAuto = d.source === 'feuilleHeure';

    if (d.phase === 'Initial') totalInitial += totalLigne;
    if (d.phase === 'En cours') totalEncours += totalLigne;
    if (d.phase === 'Reste à faire') totalReste += totalLigne;

    return `
      <div class="depense-item" style="
        border:1px solid ${estFraisGeneraux ? '#f59e0b' : estMainDoeuvreAuto ? '#10b981' : '#ccc'};
        background:${estFraisGeneraux ? '#fff7ed' : estMainDoeuvreAuto ? '#ecfdf5' : 'white'};
        padding:10px;
        margin:5px;
        border-radius:8px;
      ">
        <strong>${d.numero || ''} - ${d.nom || ''}</strong><br>
        Type : ${d.type || ''} ${estFraisGeneraux ? '⚙️' : ''}${estMainDoeuvreAuto ? ' 🕒' : ''}<br>
        Phase : ${d.phase || ''}<br>
        Fournisseur : ${d.fournisseur || ''}<br>
        Facture : ${d.facture || ''}<br>
        Date : ${d.date || ''}<br>
        Description : ${d.description || ''}<br>
        Quantité : ${d.quantite || ''}<br>
        Unité : ${d.unite || ''}<br>
        Prix unitaire : ${montant} €<br>
        <strong>Prix total : ${totalLigne} €</strong><br><br>

        ${(estFraisGeneraux || estMainDoeuvreAuto) ? `
          <span style="color:#166534;font-weight:bold;">Ligne automatique verrouillée</span>
        ` : d.id ? `
          <button onclick="chargerDepense(
            '${d.id}',
            '${d.numero || ''}',
            '${d.nom || ''}',
            '${d.type || ''}',
            '${d.phase || ''}',
            '${d.fournisseur || ''}',
            '${d.facture || ''}',
            '${d.date || ''}',
            '${d.description || ''}',
            '${d.quantite || ''}',
            '${d.unite || ''}',
            '${d.montant || ''}'
          )" style="background:orange;color:white;border:none;border-radius:5px;padding:8px;margin-right:5px;width:auto;">
            Modifier
          </button>

          <button onclick="supprimerDepense(${d.id})" style="background:red;color:white;border:none;border-radius:5px;padding:8px;width:auto;">
            Supprimer
          </button>
        ` : `
          <span style="color:#999;">Suppression indisponible</span>
        `}
      </div>
    `;
  }).join('') || '<p>Aucune dépense</p>';

  const totalGeneral = totalInitial + totalEncours + totalReste;
  const delta = totalInitial - (totalEncours + totalReste);

  const suiviParChantier = {};
  data.depenses.forEach(d => {
    const key = (d.numero || '') + ' - ' + (d.nom || '');
    const quantite = Number(d.quantite) || 0;
    const montant = Number(d.montant) || 0;
    const total = quantite * montant;

    if (!suiviParChantier[key]) suiviParChantier[key] = 0;
    suiviParChantier[key] += total;
  });

  const htmlChantiers = data.chantiers.map(c => {
    const key = (c.numero || '') + ' - ' + (c.nom || '');
    const totalDepenses = suiviParChantier[key] || 0;
    const totalMarche = Number(c.totalMarcheHT) || 0;
    const reste = totalMarche - totalDepenses;

    return `
      <div class="chantier-item" style="border:1px solid #ccc;padding:10px;margin:5px;border-radius:8px;">
        <strong>${c.numero || ''} - ${c.nom || ''}</strong><br>
        Date chantier : ${c.date || ''}<br>
        Maître d’ouvrage : ${c.maitreOuvrage || ''}<br>
        Maître d’œuvre : ${c.maitreOeuvre || ''}<br>
        Lieu : ${c.lieu || ''}<br><br>

        💰 Total marché HT : ${totalMarche} €<br>
        💸 Dépenses : ${totalDepenses} €<br>
        📉 Reste : ${reste} €<br><br>

        ${
          reste < 0
            ? `<span style="color:red;font-weight:bold;">⚠️ DÉPASSEMENT</span>`
            : `<span style="color:green;">✔ OK</span>`
        }
        <br><br>

        <a href="/export-chantier-xlsx/${c.numero}">
          <button type="button" style="background:blue;color:white;width:auto;">Exporter Excel</button>
        </a>

        <button type="button" onclick="supprimerChantier('${c.numero}')"
          style="background:red;color:white;margin-left:5px;width:auto;">
          Supprimer
        </button>
      </div>
    `;
  }).join('') || '<p>Aucun chantier</p>';

  const htmlSuiviChantier = Object.keys(suiviParChantier).map(nomChantier => `
    <div style="border:1px solid #ccc;padding:10px;margin:5px;border-radius:8px;">
      <strong>${nomChantier}</strong><br>
      Total dépenses : ${suiviParChantier[nomChantier]} €
    </div>
  `).join('') || '<p>Aucun suivi chantier.</p>';

  const htmlFeuillesHeures = data.feuillesHeures.map(fh => `
    <div style="border:1px solid #ccc;padding:10px;margin:5px;border-radius:8px;background:#f8fafc;">
      <strong>${fh.numero || ''} - ${fh.nom || ''}</strong><br>
      Nom / prénom : ${fh.personne || ''}<br>
      Rôle : ${fh.role || ''}<br>
      Date : ${fh.date || ''}<br>
      Nombre d'heure : ${fh.heures || ''}<br>
      Total : ${fh.total || 0} €<br><br>

      <button onclick="chargerFeuilleHeure(
        '${fh.id}',
        '${fh.numero || ''}',
        '${fh.nom || ''}',
        '${fh.role || ''}',
        '${fh.personne || ''}',
        '${fh.date || ''}',
        '${fh.heures || ''}'
      )" style="background:orange;color:white;border:none;border-radius:5px;padding:8px;margin-right:5px;width:auto;">
        Modifier
      </button>

      <button onclick="supprimerFeuilleHeure(${fh.id})" style="background:red;color:white;border:none;border-radius:5px;padding:8px;width:auto;">
        Supprimer
      </button>
    </div>
  `).join('') || '<p>Aucune feuille d’heure</p>';

  const interventionsTriees = [...data.interventions].sort((a, b) => {
    const dateA = new Date((a.datePrevue || '') + ' ' + (a.heurePrevue || '00:00'));
    const dateB = new Date((b.datePrevue || '') + ' ' + (b.heurePrevue || '00:00'));
    return dateA - dateB;
  });

  const planning = {};
  interventionsTriees.forEach(i => {
    const date = i.datePrevue || 'Sans date';
    if (!planning[date]) planning[date] = [];
    planning[date].push(i);
  });

  const htmlPlanning = Object.keys(planning).map(date => {
    const interventionsJour = planning[date].map(i => {
      let couleurFond = '#dbeafe';
      let couleurBord = '#60a5fa';

      if ((i.statut || '').toLowerCase() === 'terminée') {
        couleurFond = '#dcfce7';
        couleurBord = '#22c55e';
      } else if ((i.statut || '').toLowerCase() === 'en cours') {
        couleurFond = '#fed7aa';
        couleurBord = '#f97316';
      } else if ((i.statut || '').toLowerCase() === 'reportée') {
        couleurFond = '#e5e7eb';
        couleurBord = '#6b7280';
      }

      if ((i.priorite || '').toLowerCase() === 'critique') {
        couleurFond = '#fee2e2';
        couleurBord = '#ef4444';
      }

      return `
        <div 
          class="planning-item"
          data-technicien="${(i.technicien || '').toLowerCase()}"
          data-statut="${(i.statut || '').toLowerCase()}"
          data-date="${i.datePrevue || ''}"
          style="
            margin-left:20px;
            margin-top:8px;
            padding:8px;
            border-radius:8px;
            border:1px solid ${couleurBord};
            background:${couleurFond};
          ">
          ⏰ ${i.heurePrevue || ''}
          → ${i.numero || ''} - ${i.nom || ''}
          → ${i.technicien || ''}
          <br>
          <small>Statut : ${i.statut || ''} | Priorité : ${i.priorite || ''}</small>
        </div>
      `;
    }).join('');

    return `
      <div class="planning-jour" style="border:1px solid #ccc;padding:10px;margin:10px;border-radius:8px;background:#eef2ff;">
        <strong>📅 ${date}</strong>
        ${interventionsJour}
      </div>
    `;
  }).join('') || '<p>Aucune intervention planifiée</p>';

  const htmlInterventions = data.interventions.map(i => `
    <div style="border:1px solid #ccc;padding:10px;margin:5px;border-radius:8px;background:#f8fafc;">
      <strong>${i.numero || ''} - ${i.nom || ''}</strong><br>
      Client : ${i.client || ''}<br>
      Site : ${i.site || ''}<br>
      Installation : ${i.typeInstallation || ''}<br>
      Intervention : ${i.typeIntervention || ''}<br>
      Date prévue : ${i.datePrevue || ''} ${i.heurePrevue || ''}<br>
      Technicien : ${i.technicien || ''}<br>
      Priorité : ${i.priorite || ''}<br>
      Statut : ${i.statut || ''}<br><br>

      Pompe 1 heures : ${i.pompe1Heure || ''} | kWh : ${i.pompe1Kwh || ''} | État : ${i.pompe1Etat || ''}<br>
      Pompe 2 heures : ${i.pompe2Heure || ''} | kWh : ${i.pompe2Kwh || ''} | État : ${i.pompe2Etat || ''}<br>
      Panier dégrilleur : ${i.degrilleur || ''}<br>
      Flotteurs testés : ${i.flotteur || ''}<br>
      Alarme testée : ${i.alarme || ''}<br>
      Nettoyage : ${i.nettoyage || ''}<br>
      Observations : ${i.observations || ''}<br>
      Anomalies : ${i.anomalies || ''}<br><br>
      ${i.signature ? `<img src="${i.signature}" style="max-width:200px;border:1px solid #000;"><br><br>` : ''}

      <button onclick="chargerIntervention(
        '${i.id}',
        '${i.numero || ''}',
        '${i.nom || ''}',
        '${i.client || ''}',
        '${i.site || ''}',
        '${i.typeInstallation || ''}',
        '${i.typeIntervention || ''}',
        '${i.datePrevue || ''}',
        '${i.heurePrevue || ''}',
        '${i.technicien || ''}',
        '${i.priorite || ''}',
        '${i.statut || ''}',
        '${i.pompe1Heure || ''}',
        '${i.pompe1Kwh || ''}',
        '${i.pompe1Etat || ''}',
        '${i.pompe2Heure || ''}',
        '${i.pompe2Kwh || ''}',
        '${i.pompe2Etat || ''}',
        '${i.degrilleur || ''}',
        '${i.flotteur || ''}',
        '${i.alarme || ''}',
        '${i.nettoyage || ''}',
        '${i.observations || ''}',
        '${i.anomalies || ''}'
      )" style="background:orange;color:white;border:none;border-radius:5px;padding:8px;margin-right:5px;width:auto;">
        Modifier
      </button>

      <button onclick="supprimerIntervention(${i.id})" style="background:red;color:white;border:none;border-radius:5px;padding:8px;width:auto;">
        Supprimer
      </button>

      <a href="/rapport-intervention/${i.id}" target="_blank">
        <button style="background:blue;color:white;border:none;border-radius:5px;padding:8px;margin-top:5px;width:auto;">
          Voir rapport
        </button>
      </a>

      <a href="/export-word-intervention/${i.id}" target="_blank">
        <button style="background:green;color:white;border:none;border-radius:5px;padding:8px;margin-top:5px;width:auto;">
          Exporter Word
        </button>
      </a>
    </div>
  `).join('') || '<p>Aucune intervention</p>';

  const htmlUsers = data.utilisateurs.map(u => `
    <div style="border:1px solid #ccc;padding:10px;margin:5px;border-radius:8px;">
      <strong>${u.nom || ''}</strong> - ${u.role || ''}<br>
      Droits :
      ${u.droits?.chantiers ? 'Chantiers ' : ''}
      ${u.droits?.depenses ? 'Dépenses ' : ''}
      ${u.droits?.feuillesHeures ? 'FeuillesHeures ' : ''}
      ${u.droits?.interventions ? 'Interventions ' : ''}
      ${u.droits?.photos ? 'Photos ' : ''}
      ${u.droits?.reglages ? 'Réglages ' : ''}
    </div>
  `).join('') || '<p>Aucun utilisateur</p>';

const htmlArchives = (data.archives || []).map(a => `
  <div style="border:1px solid #ccc;padding:10px;margin:10px;border-radius:8px;background:#f1f5f9;">
    <strong>Chantier :</strong> ${a.numero || ''}<br>
    <strong>Date archivage :</strong> ${new Date(a.dateArchivage).toLocaleString()}<br><br>

    📁 Dépenses : ${(a.depenses || []).length}<br>
    🕒 Feuilles d’heure : ${(a.feuillesHeures || []).length}<br>
    🔧 Interventions : ${(a.interventions || []).length}<br>
    📸 Photos : ${(a.photos || []).length}<br><br>

    <a href="/archive/${a.id}" target="_blank">
      <button style="background:blue;color:white;">Voir dossier</button>
    </a>
    <button onclick="restaurerArchive('${a.id}')" style="background:green;color:white;margin-left:5px;">
  Restaurer
</button>
  </div>
`).join('') || '<p>Aucune archive</p>';

  const htmlPhotos = data.photos.map(p => `
  <div style="border:1px solid #ccc;padding:10px;margin:10px 0;border-radius:8px;">
    <strong>📁 Chantier :</strong> ${p.chantier || 'Non défini'}<br>
    <strong>📝 Description :</strong> ${p.description}<br><br>
    ${p.fichier ? `<img src="${p.fichier}" style="max-width:300px;max-height:220px;border-radius:8px;border:1px solid #ddd;">` : '<p>Aucune image</p>'}
  </div>
`).join('') || '<p>Aucune photo</p>';

  const blocReglages = droits.reglages ? `
    <div class="bloc">
      <h3>Réglages</h3>
      <p>Pourcentage frais généraux automatique actuel : <strong>${pourcentageFraisGeneraux}%</strong></p>
      <input id="r_frais_generaux" placeholder="Pourcentage frais généraux" value="${pourcentageFraisGeneraux}">
      <button onclick="enregistrerReglageFraisGeneraux()">Enregistrer le pourcentage</button>

      <p style="margin-top:20px;">Prix heure Cadre actuel : <strong>${prixHeureCadre} €</strong></p>
      <input id="r_prix_cadre" placeholder="Prix heure Cadre" value="${prixHeureCadre}">
      <button onclick="enregistrerPrixCadre()">Enregistrer prix Cadre</button>

      <p style="margin-top:20px;">Prix heure Technicien actuel : <strong>${prixHeureTechnicien} €</strong></p>
      <input id="r_prix_technicien" placeholder="Prix heure Technicien" value="${prixHeureTechnicien}">
      <button onclick="enregistrerPrixTechnicien()">Enregistrer prix Technicien</button>

      <hr>

      <h3>Utilisateurs</h3>
      <input id="u_nom" placeholder="Nom">
      <input id="u_role" placeholder="Rôle">
      <input id="u_motdepasse" placeholder="Mot de passe">

      <br><br>

      <label><input type="checkbox" id="u_chantiers"> Chantiers</label>
      <label><input type="checkbox" id="u_depenses"> Dépenses</label>
      <label><input type="checkbox" id="u_feuillesHeures"> Feuilles d'heure</label>
      <label><input type="checkbox" id="u_interventions"> Interventions</label>
      <label><input type="checkbox" id="u_photos"> Photos</label>
      <label><input type="checkbox" id="u_reglages"> Réglages</label>

      <br><br>

      <button onclick="addUser()">Ajouter</button>

      <div style="margin-top:15px;">
        ${htmlUsers}
      </div>

<hr>

<h3>Archives</h3>

<div style="margin-top:15px;">
  ${htmlArchives}
</div>
    </div>
  ` : '';

  const blocPlanning = droits.interventions ? `
    <div class="bloc">
      <h3>Planning des interventions</h3>

      <p>
        <a href="/planning-pdf" target="_blank">
          <button type="button" style="background:#dc2626;color:white;width:auto;">
            Export PDF planning
          </button>
        </a>
      </p>

      <div style="margin-bottom:10px;">
        <input id="filtre_technicien_planning" placeholder="Filtrer par technicien">

        <select id="filtre_statut_planning">
          <option value="">Tous les statuts</option>
          <option value="à planifier">à planifier</option>
          <option value="planifiée">planifiée</option>
          <option value="en cours">en cours</option>
          <option value="terminée">terminée</option>
          <option value="reportée">reportée</option>
        </select>

        <input id="filtre_date_planning" type="date">

        <button onclick="filtrerPlanningComplet()">Filtrer</button>
        <button onclick="resetPlanningComplet()">Réinitialiser</button>
      </div>

      ${htmlPlanning}
    </div>
  ` : '';

  const blocChantiers = droits.chantiers ? `
    <div class="bloc" id="chantiers">
      <h3>Chantiers</h3>
      <input id="c_numero" placeholder="Numéro chantier">
      <input id="c_nom" placeholder="Nom chantier">
      <input id="c_date" type="date">
      <input id="c_maitreOuvrage" placeholder="Maître d’ouvrage">
      <input id="c_maitreOeuvre" placeholder="Maître d’œuvre">
      <input id="c_lieu" placeholder="Lieu">
      <input id="c_totalMarcheHT" placeholder="Total marché HT">
      <button onclick="addChantier()">Ajouter</button>

      <div style="margin:10px 0;">
        <input id="filtre_chantier_liste" placeholder="Filtrer chantier">
        <button onclick="filtrerListeChantier()">Filtrer</button>
        <button onclick="resetChantier()">Reset</button>
      </div>

      ${htmlChantiers}
    </div>
  ` : '';

  const blocDepenses = droits.depenses ? `
    <div class="bloc" id="depenses">
      <h3>Dépenses chantier</h3>
      <p><strong>Initial :</strong> ${totalInitial} €</p>
      <p><strong>En cours :</strong> ${totalEncours} €</p>
      <p><strong>Reste à faire :</strong> ${totalReste} €</p>
      <p><strong>Total général :</strong> ${totalGeneral} €</p>
      <p><strong>Delta :</strong> ${delta} €</p>
      <p>
        <a href="/export-xlsx">
          <button type="button">Exporter en vrai Excel XLSX</button>
        </a>
      </p>
      <p>
        <a href="/export-excel">
          <button type="button">Exporter en Excel</button>
        </a>
      </p>

      <input id="d_id" type="hidden">
      <select id="d_numero" onchange="remplirNomChantier()">
        <option value="">Choisir un chantier</option>
        ${optionsChantiers}
      </select>

      <input id="d_nom" placeholder="Nom chantier" readonly>

      <select id="d_type">
        <option>Fourniture</option>
        <option>Main d'oeuvre</option>
        <option>Sous-traitance</option>
        <option>Étude</option>
      </select>

      <select id="d_phase">
        <option>Initial</option>
        <option>En cours</option>
        <option>Reste à faire</option>
      </select>

      <input id="d_fournisseur" placeholder="Fournisseur">
      <input id="d_facture" placeholder="Numéro facture">
      <input id="d_date" type="date">
      <input id="d_description" placeholder="Description">
      <input id="d_quantite" placeholder="Quantité" oninput="calculerTotalDepense()">

      <select id="d_unite">
        <option value="">Unité</option>
        <option>M</option>
        <option>M2</option>
        <option>U</option>
        <option>H</option>
      </select>

      <input id="d_montant" placeholder="Prix unitaire" oninput="calculerTotalDepense()">
      <input id="d_total" placeholder="Prix total" readonly>

      <button onclick="addDepense()">Ajouter</button>
      <button onclick="modifierDepense()" style="background:green;color:white;">
        Enregistrer modification
      </button>

      <div style="margin:10px 0;">
        <input id="filtre_chantier" placeholder="Filtrer par numéro ou nom chantier">
        <button onclick="filtrerChantier()">Filtrer</button>
        <button onclick="reinitialiserFiltre()">Réinitialiser</button>
      </div>

      ${htmlDepenses}
    </div>
  ` : '';

  const blocFeuillesHeures = droits.feuillesHeures ? `
    <div class="bloc">
      <h3>Feuille d’heure</h3>
      <p>Prix heure Cadre : <strong>${prixHeureCadre} €</strong> | Prix heure Technicien : <strong>${prixHeureTechnicien} €</strong></p>

      <input id="fh_id" type="hidden">

      <select id="fh_numero" onchange="remplirNomChantierFeuille()">
        <option value="">Choisir un chantier</option>
        ${optionsChantiers}
      </select>

      <input id="fh_nom" placeholder="Nom chantier" readonly>

      <select id="fh_role" onchange="calculerTotalFeuilleHeure()">
        <option value="Technicien">Technicien</option>
        <option value="Cadre">Cadre</option>
      </select>

      <input id="fh_personne" placeholder="Nom et prénom">
      <input id="fh_date" type="date">
      <input id="fh_heures" placeholder="Nombre d'heure" oninput="calculerTotalFeuilleHeure()">
      <input id="fh_total" placeholder="Total" readonly>

      <button onclick="addFeuilleHeure()">Ajouter feuille d’heure</button>
      <button onclick="modifierFeuilleHeure()" style="background:green;color:white;">Modifier feuille d’heure</button>

      <p style="margin-top:15px;">
        <a href="/export-feuilles-heures-xlsx">
          <button type="button">Exporter Excel feuilles d’heure</button>
        </a>
      </p>

      ${htmlFeuillesHeures}
    </div>
  ` : '';

  const blocInterventions = droits.interventions ? `
    <div class="bloc" id="interventions">
      <h3>Interventions</h3>

      <input id="i_id" type="hidden">

      <select id="i_numero" onchange="remplirNomChantierIntervention()">
        <option value="">Choisir un chantier</option>
        ${optionsChantiers}
      </select>

      <input id="i_nom" placeholder="Nom chantier" readonly>
      <input id="i_client" placeholder="Client">
      <input id="i_site" placeholder="Site">

      <select id="i_typeInstallation">
        <option>citerne eau potable</option>
        <option>citerne eau de pluie</option>
        <option>poste de relevage</option>
        <option>station traitement eau potable</option>
        <option>station d'épuration</option>
      </select>

      <select id="i_typeIntervention">
        <option>entretien</option>
        <option>dépannage</option>
        <option>contrôle</option>
        <option>nettoyage</option>
        <option>urgence</option>
      </select>

      <input id="i_datePrevue" type="date">
      <input id="i_heurePrevue" type="time">
      <input id="i_technicien" placeholder="Technicien">

      <select id="i_priorite">
        <option>normale</option>
        <option>urgente</option>
        <option>critique</option>
      </select>

      <select id="i_statut">
        <option>à planifier</option>
        <option>planifiée</option>
        <option>en cours</option>
        <option>terminée</option>
        <option>reportée</option>
      </select>

      <h4>Pompe 1</h4>
      <input id="i_pompe1Heure" placeholder="Pompe 1 heures">
      <input id="i_pompe1Kwh" placeholder="Pompe 1 kWh">
      <input id="i_pompe1Etat" placeholder="Pompe 1 état">

      <h4>Pompe 2</h4>
      <input id="i_pompe2Heure" placeholder="Pompe 2 heures">
      <input id="i_pompe2Kwh" placeholder="Pompe 2 kWh">
      <input id="i_pompe2Etat" placeholder="Pompe 2 état">

      <select id="i_degrilleur">
        <option value="">Panier dégrilleur nettoyé ?</option>
        <option>Oui</option>
        <option>Non</option>
      </select>

      <select id="i_flotteur">
        <option value="">Flotteurs testés ?</option>
        <option>Oui</option>
        <option>Non</option>
      </select>

      <select id="i_alarme">
        <option value="">Alarme testée ?</option>
        <option>Oui</option>
        <option>Non</option>
      </select>

      <select id="i_nettoyage">
        <option value="">Nettoyage effectué ?</option>
        <option>Oui</option>
        <option>Non</option>
      </select>

      <textarea id="i_observations" placeholder="Observations"></textarea>
      <textarea id="i_anomalies" placeholder="Anomalies"></textarea>

      <h4>Signature client</h4>
      <canvas id="signaturePad" width="300" height="150" style="border:1px solid #000;background:white;"></canvas><br>
      <button type="button" onclick="clearSignature()" style="background:gray;color:white;">Effacer signature</button>

      <input id="i_signature" type="hidden">

      <button onclick="addIntervention()">Ajouter intervention</button>
      <button onclick="modifierIntervention()" style="background:green;color:white;">Modifier intervention</button>

      <p style="margin-top:15px;">
        <a href="/export-interventions-xlsx">
          <button type="button">Exporter Excel interventions</button>
        </a>
      </p>

      ${htmlInterventions}
    </div>
  ` : '';

  const blocSuivi = droits.depenses ? `
    <div class="bloc">
      <h3>Suivi par chantier</h3>
      ${htmlSuiviChantier}
    </div>
  ` : '';

  const blocPhotos = droits.photos ? `
    <div class="bloc" id="photos">
      <h3>Photos chantier</h3>
      <form action="/photo" method="POST" enctype="multipart/form-data">
        <select name="chantier" required>
  <option value="">Choisir un chantier</option>
  ${optionsChantiers}
</select>
        <input name="description" placeholder="Description photo" required>
        <input type="file" name="photo" accept="image/*" required>
        <button type="submit">Envoyer la photo</button>
      </form>
      ${htmlPhotos}
    </div>
  ` : '';

  return `
    <html>
    <head>
      <meta charset="utf-8">
      <title>HDS Interne</title>
      <style>
        body { font-family: Arial; background:#f3f4f6; margin:0; padding:0; }
        .top {
          background:white;
          padding:10px 15px;
          display:flex;
          justify-content:space-between;
          align-items:center;
        }
        .bloc {
          background:white;
          padding:20px;
          margin:20px;
          border-radius:12px;
          box-shadow:0 2px 8px rgba(0,0,0,0.1);
        }
        input, button, select, textarea {
  margin:5px;
  padding:12px;
  width:100%;
  font-size:16px;
  box-sizing:border-box;
}

input, select, textarea {
  border:1px solid #ccc;
  border-radius:8px;
}

textarea {
  width:100%;
  min-height:80px;
}

button {
  width:100%;
  font-size:16px;
}

button:hover { opacity:0.9; }

.logout { background:#dc2626; }

.message-ok {
  background:#dcfce7;
  color:#166534;
  padding:10px;
  margin:15px 10px;
  border-radius:8px;
  border:1px solid #bbf7d0;
}
        .message-ko {
          background:#fee2e2;
          color:#991b1b;
          padding:10px;
          margin:15px 20px;
          border-radius:8px;
          border:1px solid #fecaca;
        }
      </style>
    </head>
    <body>
      <div style="text-align:center; background:#1e293b; padding:15px;">
        <img src="/uploads/logo.jpg" style="height:60px;"><br>
        <span style="color:white;font-size:20px;">HDS Interne</span>
      </div>

      ${message ? `<div class="${message.startsWith('ERREUR') ? 'message-ko' : 'message-ok'}">${message}</div>` : ''}

      <div class="top">
        <div>Connecté : <strong>${user.nom}</strong> (${user.role})</div>
        <form method="POST" action="/logout" style="margin:0;">
          <button class="logout" type="submit">Se déconnecter</button>
        </form>
      </div>

      ${blocReglages}
      ${blocPlanning}
      ${blocChantiers}
      ${blocDepenses}
      ${blocFeuillesHeures}
      ${blocInterventions}
      ${blocSuivi}
      ${blocPhotos}

      <script>
      function showBloc(id) {
  const blocs = document.querySelectorAll('.bloc');
  blocs.forEach(b => b.style.display = 'none');

  const actif = document.getElementById(id);
  if (actif) actif.style.display = 'block';
}

        function send(url, data) {
          fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function remplirNomChantier() {
          const select = document.getElementById('d_numero');
          if (!select) return;
          const option = select.options[select.selectedIndex];
          const nom = option ? (option.getAttribute('data-nom') || '') : '';
          const champ = document.getElementById('d_nom');
          if (champ) champ.value = nom;
        }

        function remplirNomChantierFeuille() {
          const select = document.getElementById('fh_numero');
          if (!select) return;
          const option = select.options[select.selectedIndex];
          const nom = option ? (option.getAttribute('data-nom') || '') : '';
          const champ = document.getElementById('fh_nom');
          if (champ) champ.value = nom;
        }

        function remplirNomChantierIntervention() {
          const select = document.getElementById('i_numero');
          if (!select) return;
          const option = select.options[select.selectedIndex];
          const nom = option ? (option.getAttribute('data-nom') || '') : '';
          const champ = document.getElementById('i_nom');
          if (champ) champ.value = nom;
        }

        function chargerDepense(id, numero, nom, type, phase, fournisseur, facture, date, description, quantite, unite, montant) {
          document.getElementById('d_id').value = id;
          document.getElementById('d_numero').value = numero;
          remplirNomChantier();
          document.getElementById('d_type').value = type;
          document.getElementById('d_phase').value = phase;
          document.getElementById('d_fournisseur').value = fournisseur;
          document.getElementById('d_facture').value = facture;
          document.getElementById('d_date').value = date;
          document.getElementById('d_description').value = description;
          document.getElementById('d_quantite').value = quantite;
          document.getElementById('d_unite').value = unite;
          document.getElementById('d_montant').value = montant;
          calculerTotalDepense();
        }

        function modifierDepense() {
          const id = document.getElementById('d_id').value;
          if (!id) {
            alert('Choisis une dépense à modifier');
            return;
          }

          fetch('/modifier-depense', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              id: id,
              numero: document.getElementById('d_numero').value,
              nom: document.getElementById('d_nom').value,
              type: document.getElementById('d_type').value,
              phase: document.getElementById('d_phase').value,
              fournisseur: document.getElementById('d_fournisseur').value,
              facture: document.getElementById('d_facture').value,
              date: document.getElementById('d_date').value,
              description: document.getElementById('d_description').value,
              quantite: document.getElementById('d_quantite').value,
              unite: document.getElementById('d_unite').value,
              montant: document.getElementById('d_montant').value
            })
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function supprimerDepense(id) {
          if (!confirm('Supprimer cette dépense ?')) return;

          fetch('/supprimer-depense', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id })
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function addChantier() {
          send('/chantier', {
            numero: document.getElementById('c_numero').value,
            nom: document.getElementById('c_nom').value,
            date: document.getElementById('c_date').value,
            maitreOuvrage: document.getElementById('c_maitreOuvrage').value,
            maitreOeuvre: document.getElementById('c_maitreOeuvre').value,
            lieu: document.getElementById('c_lieu').value,
            totalMarcheHT: document.getElementById('c_totalMarcheHT').value
          });
        }

        function supprimerChantier(numero) {
  const archiver = confirm('Voulez-vous archiver ce chantier ?');

  if (!archiver) {
    const confirmerSuppression = confirm('Si vous continuez, le chantier sera supprimé sans archivage. Confirmer ?');
    if (!confirmerSuppression) return;
  }

  fetch('/supprimer-chantier', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ numero, archiver })
  }).then(response => {
    if (!response.ok) {
      response.text().then(msg => alert(msg));
      return;
    }
    location.reload();
  });
}
        function filtrerListeChantier() {
          const filtre = document.getElementById('filtre_chantier_liste').value.toLowerCase();
          const blocs = document.querySelectorAll('.chantier-item');

          blocs.forEach(b => {
            const texte = b.innerText.toLowerCase();
            b.style.display = texte.includes(filtre) ? 'block' : 'none';
          });
        }

        function resetChantier() {
          document.getElementById('filtre_chantier_liste').value = '';
          document.querySelectorAll('.chantier-item').forEach(b => b.style.display = 'block');
        }

        function filtrerChantier() {
          const filtre = document.getElementById('filtre_chantier').value.toLowerCase();
          const lignes = document.querySelectorAll('.depense-item');

          lignes.forEach(ligne => {
            const texte = ligne.innerText.toLowerCase();
            ligne.style.display = texte.includes(filtre) ? 'block' : 'none';
          });
        }

        function reinitialiserFiltre() {
          document.getElementById('filtre_chantier').value = '';
          document.querySelectorAll('.depense-item').forEach(ligne => {
            ligne.style.display = 'block';
          });
        }

        function calculerTotalDepense() {
          const quantite = Number(document.getElementById('d_quantite').value) || 0;
          const montant = Number(document.getElementById('d_montant').value) || 0;
          document.getElementById('d_total').value = quantite * montant;
        }

        function addDepense() {
          send('/depense', {
            numero: document.getElementById('d_numero').value,
            nom: document.getElementById('d_nom').value,
            type: document.getElementById('d_type').value,
            phase: document.getElementById('d_phase').value,
            fournisseur: document.getElementById('d_fournisseur').value,
            facture: document.getElementById('d_facture').value,
            date: document.getElementById('d_date').value,
            description: document.getElementById('d_description').value,
            quantite: document.getElementById('d_quantite').value,
            unite: document.getElementById('d_unite').value,
            montant: document.getElementById('d_montant').value
          });
        }

        function calculerTotalFeuilleHeure() {
          const heures = Number(document.getElementById('fh_heures').value) || 0;
          const role = document.getElementById('fh_role').value;
          const prixCadre = Number(document.getElementById('r_prix_cadre') ? document.getElementById('r_prix_cadre').value : ${prixHeureCadre}) || ${prixHeureCadre};
          const prixTechnicien = Number(document.getElementById('r_prix_technicien') ? document.getElementById('r_prix_technicien').value : ${prixHeureTechnicien}) || ${prixHeureTechnicien};
          const prixHeure = role === 'Cadre' ? prixCadre : prixTechnicien;

          document.getElementById('fh_total').value = heures * prixHeure;
        }

        function addFeuilleHeure() {
          send('/feuille-heure', {
            numero: document.getElementById('fh_numero').value,
            nom: document.getElementById('fh_nom').value,
            role: document.getElementById('fh_role').value,
            personne: document.getElementById('fh_personne').value,
            date: document.getElementById('fh_date').value,
            heures: document.getElementById('fh_heures').value
          });
        }

        function chargerFeuilleHeure(id, numero, nom, role, personne, date, heures) {
          document.getElementById('fh_id').value = id;
          document.getElementById('fh_numero').value = numero;
          remplirNomChantierFeuille();
          document.getElementById('fh_role').value = role;
          document.getElementById('fh_personne').value = personne;
          document.getElementById('fh_date').value = date;
          document.getElementById('fh_heures').value = heures;
          calculerTotalFeuilleHeure();
        }

        function modifierFeuilleHeure() {
          const id = document.getElementById('fh_id').value;
          if (!id) {
            alert('Choisis une feuille d’heure à modifier');
            return;
          }

          fetch('/modifier-feuille-heure', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              id: id,
              numero: document.getElementById('fh_numero').value,
              nom: document.getElementById('fh_nom').value,
              role: document.getElementById('fh_role').value,
              personne: document.getElementById('fh_personne').value,
              date: document.getElementById('fh_date').value,
              heures: document.getElementById('fh_heures').value
            })
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function supprimerFeuilleHeure(id) {
          if (!confirm('Supprimer cette feuille d’heure ?')) return;

          fetch('/supprimer-feuille-heure', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id })
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function addIntervention() {
          const canvas = document.getElementById('signaturePad');
          document.getElementById('i_signature').value = canvas.toDataURL('image/png');

          send('/intervention', {
            numero: document.getElementById('i_numero').value,
            nom: document.getElementById('i_nom').value,
            client: document.getElementById('i_client').value,
            site: document.getElementById('i_site').value,
            typeInstallation: document.getElementById('i_typeInstallation').value,
            typeIntervention: document.getElementById('i_typeIntervention').value,
            datePrevue: document.getElementById('i_datePrevue').value,
            heurePrevue: document.getElementById('i_heurePrevue').value,
            technicien: document.getElementById('i_technicien').value,
            priorite: document.getElementById('i_priorite').value,
            statut: document.getElementById('i_statut').value,
            pompe1Heure: document.getElementById('i_pompe1Heure').value,
            pompe1Kwh: document.getElementById('i_pompe1Kwh').value,
            pompe1Etat: document.getElementById('i_pompe1Etat').value,
            pompe2Heure: document.getElementById('i_pompe2Heure').value,
            pompe2Kwh: document.getElementById('i_pompe2Kwh').value,
            pompe2Etat: document.getElementById('i_pompe2Etat').value,
            degrilleur: document.getElementById('i_degrilleur').value,
            flotteur: document.getElementById('i_flotteur').value,
            alarme: document.getElementById('i_alarme').value,
            nettoyage: document.getElementById('i_nettoyage').value,
            observations: document.getElementById('i_observations').value,
            anomalies: document.getElementById('i_anomalies').value,
            signature: document.getElementById('i_signature').value
          });
        }

        function chargerIntervention(id, numero, nom, client, site, typeInstallation, typeIntervention, datePrevue, heurePrevue, technicien, priorite, statut, pompe1Heure, pompe1Kwh, pompe1Etat, pompe2Heure, pompe2Kwh, pompe2Etat, degrilleur, flotteur, alarme, nettoyage, observations, anomalies) {
          document.getElementById('i_id').value = id;
          document.getElementById('i_numero').value = numero;
          remplirNomChantierIntervention();
          document.getElementById('i_client').value = client;
          document.getElementById('i_site').value = site;
          document.getElementById('i_typeInstallation').value = typeInstallation;
          document.getElementById('i_typeIntervention').value = typeIntervention;
          document.getElementById('i_datePrevue').value = datePrevue;
          document.getElementById('i_heurePrevue').value = heurePrevue;
          document.getElementById('i_technicien').value = technicien;
          document.getElementById('i_priorite').value = priorite;
          document.getElementById('i_statut').value = statut;
          document.getElementById('i_pompe1Heure').value = pompe1Heure;
          document.getElementById('i_pompe1Kwh').value = pompe1Kwh;
          document.getElementById('i_pompe1Etat').value = pompe1Etat;
          document.getElementById('i_pompe2Heure').value = pompe2Heure;
          document.getElementById('i_pompe2Kwh').value = pompe2Kwh;
          document.getElementById('i_pompe2Etat').value = pompe2Etat;
          document.getElementById('i_degrilleur').value = degrilleur;
          document.getElementById('i_flotteur').value = flotteur;
          document.getElementById('i_alarme').value = alarme;
          document.getElementById('i_nettoyage').value = nettoyage;
          document.getElementById('i_observations').value = observations;
          document.getElementById('i_anomalies').value = anomalies;
        }

        function modifierIntervention() {
          const canvas = document.getElementById('signaturePad');
          document.getElementById('i_signature').value = canvas.toDataURL('image/png');

          const id = document.getElementById('i_id').value;
          if (!id) {
            alert('Choisis une intervention à modifier');
            return;
          }

          fetch('/modifier-intervention', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              id: id,
              numero: document.getElementById('i_numero').value,
              nom: document.getElementById('i_nom').value,
              client: document.getElementById('i_client').value,
              site: document.getElementById('i_site').value,
              typeInstallation: document.getElementById('i_typeInstallation').value,
              typeIntervention: document.getElementById('i_typeIntervention').value,
              datePrevue: document.getElementById('i_datePrevue').value,
              heurePrevue: document.getElementById('i_heurePrevue').value,
              technicien: document.getElementById('i_technicien').value,
              priorite: document.getElementById('i_priorite').value,
              statut: document.getElementById('i_statut').value,
              pompe1Heure: document.getElementById('i_pompe1Heure').value,
              pompe1Kwh: document.getElementById('i_pompe1Kwh').value,
              pompe1Etat: document.getElementById('i_pompe1Etat').value,
              pompe2Heure: document.getElementById('i_pompe2Heure').value,
              pompe2Kwh: document.getElementById('i_pompe2Kwh').value,
              pompe2Etat: document.getElementById('i_pompe2Etat').value,
              degrilleur: document.getElementById('i_degrilleur').value,
              flotteur: document.getElementById('i_flotteur').value,
              alarme: document.getElementById('i_alarme').value,
              nettoyage: document.getElementById('i_nettoyage').value,
              observations: document.getElementById('i_observations').value,
              anomalies: document.getElementById('i_anomalies').value,
              signature: document.getElementById('i_signature').value
            })
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function supprimerIntervention(id) {
          if (!confirm('Supprimer cette intervention ?')) return;

          fetch('/supprimer-intervention', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id })
          }).then(response => {
            if (!response.ok) {
              response.text().then(msg => alert(msg));
              return;
            }
            location.reload();
          });
        }

        function addUser() {
          send('/utilisateur', {
            nom: document.getElementById('u_nom').value,
            role: document.getElementById('u_role').value,
            motdepasse: document.getElementById('u_motdepasse').value,
            chantiers: document.getElementById('u_chantiers').checked,
            depenses: document.getElementById('u_depenses').checked,
            feuillesHeures: document.getElementById('u_feuillesHeures').checked,
            interventions: document.getElementById('u_interventions').checked,
            photos: document.getElementById('u_photos').checked,
            reglages: document.getElementById('u_reglages').checked
          });
        }

        function enregistrerReglageFraisGeneraux() {
          const pourcentage = document.getElementById('r_frais_generaux').value;
          fetch('/reglages-frais-generaux', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pourcentage })
          }).then(response => {
            if (!response.ok) {
              alert('Erreur réglage');
              return;
            }
            location.reload();
          });
        }

        function enregistrerPrixCadre() {
          const prix = document.getElementById('r_prix_cadre').value;
          fetch('/reglages-prix-cadre', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ prix })
          }).then(response => {
            if (!response.ok) {
              alert('Erreur prix cadre');
              return;
            }
            location.reload();
          });
        }

        function enregistrerPrixTechnicien() {
          const prix = document.getElementById('r_prix_technicien').value;
          fetch('/reglages-prix-technicien', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ prix })
          }).then(response => {
            if (!response.ok) {
              alert('Erreur prix technicien');
              return;
            }
            location.reload();
          });
        }

        const canvas = document.getElementById('signaturePad');
        if (canvas) {
          const ctx = canvas.getContext('2d');
          let dessinEnCours = false;

          canvas.addEventListener('mousedown', function () {
            dessinEnCours = true;
          });

          canvas.addEventListener('mouseup', function () {
            dessinEnCours = false;
            ctx.beginPath();
            document.getElementById('i_signature').value = canvas.toDataURL();
          });

          canvas.addEventListener('mousemove', function (e) {
            if (!dessinEnCours) return;

            ctx.lineWidth = 2;
            ctx.lineCap = 'round';
            ctx.strokeStyle = '#000';

            ctx.lineTo(e.offsetX, e.offsetY);
            ctx.stroke();
            ctx.beginPath();
            ctx.moveTo(e.offsetX, e.offsetY);
          });

          window.clearSignature = function () {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.beginPath();
            document.getElementById('i_signature').value = '';
          };
        }

        function filtrerPlanningComplet() {
          const filtreTechnicien = document.getElementById('filtre_technicien_planning').value.toLowerCase();
          const filtreStatut = document.getElementById('filtre_statut_planning').value.toLowerCase();
          const filtreDate = document.getElementById('filtre_date_planning').value;

          const jours = document.querySelectorAll('.planning-jour');

          jours.forEach(jour => {
            const lignes = jour.querySelectorAll('.planning-item');
            let auMoinsUneVisible = false;

            lignes.forEach(ligne => {
              const technicien = ligne.getAttribute('data-technicien') || '';
              const statut = ligne.getAttribute('data-statut') || '';
              const date = ligne.getAttribute('data-date') || '';

              const okTechnicien = !filtreTechnicien || technicien.includes(filtreTechnicien);
              const okStatut = !filtreStatut || statut === filtreStatut;
              const okDate = !filtreDate || date === filtreDate;

              if (okTechnicien && okStatut && okDate) {
                ligne.style.display = 'block';
                auMoinsUneVisible = true;
              } else {
                ligne.style.display = 'none';
              }
            });

            jour.style.display = auMoinsUneVisible ? 'block' : 'none';
          });
        }

        function resetPlanningComplet() {
          document.getElementById('filtre_technicien_planning').value = '';
          document.getElementById('filtre_statut_planning').value = '';
          document.getElementById('filtre_date_planning').value = '';

          const jours = document.querySelectorAll('.planning-jour');

          jours.forEach(jour => {
            jour.style.display = 'block';
            const lignes = jour.querySelectorAll('.planning-item');
            lignes.forEach(ligne => {
              ligne.style.display = 'block';
            });
          });
        }

      function restaurerArchive(id) {
  if (!confirm("Restaurer ce chantier ?")) return;

  fetch('/restaurer-archive', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ id })
  }).then(() => location.reload());
}

function showBloc(id) {
  const blocs = document.querySelectorAll('.bloc');

  blocs.forEach(b => {
    b.style.display = 'none';
  });

  const actif = document.getElementById(id);
  if (actif) {
    actif.style.display = 'block';
  }
}
      </script>
    </body>
    </html>
  `;
}

function verifierConnexion(req, res, next) {
  if (!req.session.user) {
    res.redirect('/');
    return;
  }
  next();
}

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

app.post('/login', (req, res) => {
  const data = lireDonnees();
  const { nom, motdepasse } = req.body;

  const user = data.utilisateurs.find(
    u => (u.nom || '').toLowerCase() === (nom || '').toLowerCase()
      && u.motdepasse === motdepasse
  );

  if (!user) {
    res.send(pageConnexion('Nom ou mot de passe incorrect'));
    return;
  }

  req.session.user = user;
  res.redirect('/');
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/');
  });
});

app.post('/reglages-frais-generaux', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const pourcentage = Number(req.body.pourcentage);

  if (isNaN(pourcentage) || pourcentage < 0) {
    res.status(400).send('Pourcentage invalide');
    return;
  }

  data.reglages.pourcentageFraisGeneraux = pourcentage;
  sauvegarder(data);
  res.send('OK');
});

app.post('/reglages-prix-cadre', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const prix = Number(req.body.prix);

  if (isNaN(prix) || prix < 0) {
    res.status(400).send('Prix cadre invalide');
    return;
  }

  data.reglages.prixHeureCadre = prix;
  sauvegarder(data);
  res.send('OK');
});

app.post('/reglages-prix-technicien', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const prix = Number(req.body.prix);

  if (isNaN(prix) || prix < 0) {
    res.status(400).send('Prix technicien invalide');
    return;
  }

  data.reglages.prixHeureTechnicien = prix;
  sauvegarder(data);
  res.send('OK');
});

app.post('/chantier', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const numero = (req.body.numero || '').toString().trim();
  const archiver = req.body.archiver === true;
  const nom = req.body.nom || '';
  const date = req.body.date || '';
  const maitreOuvrage = req.body.maitreOuvrage || '';
  const maitreOeuvre = req.body.maitreOeuvre || '';
  const lieu = req.body.lieu || '';
  const totalMarcheHT = Number(req.body.totalMarcheHT) || 0;

  if (!numero) {
    res.status(400).send('Numéro chantier obligatoire');
    return;
  }

  const existeDeja = data.chantiers.some(c => ((c.numero || '').toString().trim()) === numero);

  if (existeDeja) {
    res.status(400).send('Numéro chantier déjà existant');
    return;
  }

  data.chantiers.push({
    numero,
    nom,
    date,
    maitreOuvrage,
    maitreOeuvre,
    lieu,
    totalMarcheHT
  });

  const pourcentage = Number(data.reglages?.pourcentageFraisGeneraux) || 8;
  const fraisGeneraux = totalMarcheHT * (pourcentage / 100);

  data.depenses.push({
    id: Date.now(),
    numero: numero,
    nom: nom,
    type: 'Frais généraux',
    phase: 'Initial',
    fournisseur: 'Interne',
    facture: '',
    date: date,
    description: `Frais généraux automatiques ${pourcentage}%`,
    quantite: '1',
    unite: 'U',
    montant: fraisGeneraux.toString(),
    verrouillee: true
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/supprimer-chantier', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const numero = (req.body.numero || '').toString().trim();
  const archiver = req.body.archiver === true;

  if (!numero) {
    res.status(400).send('Numéro chantier manquant');
    return;
  }

  const chantier = data.chantiers.find(c => ((c.numero || '').toString().trim()) === numero);

  if (!chantier) {
    res.status(404).send('Chantier introuvable');
    return;
  }

  if (archiver) {
    const depensesArchivees = data.depenses.filter(d => ((d.numero || '').toString().trim()) === numero);
    const feuillesHeuresArchivees = data.feuillesHeures.filter(fh => ((fh.numero || '').toString().trim()) === numero);
    const interventionsArchivees = data.interventions.filter(i => ((i.numero || '').toString().trim()) === numero);
    const photosArchivees = data.photos.filter(p => (p.chantier || '').includes(numero));

    data.archives.push({
      id: Date.now(),
      dateArchivage: new Date().toISOString(),
      numero: numero,
      chantier: chantier,
      depenses: depensesArchivees,
      feuillesHeures: feuillesHeuresArchivees,
      interventions: interventionsArchivees,
      photos: photosArchivees
    });
  }

  data.chantiers = data.chantiers.filter(c => ((c.numero || '').toString().trim()) !== numero);
  data.depenses = data.depenses.filter(d => ((d.numero || '').toString().trim()) !== numero);
  data.feuillesHeures = data.feuillesHeures.filter(fh => ((fh.numero || '').toString().trim()) !== numero);
  data.interventions = data.interventions.filter(i => ((i.numero || '').toString().trim()) !== numero);
  data.photos = data.photos.filter(p => !(p.chantier || '').includes(numero));

  sauvegarder(data);
  res.send('OK');
});

app.post('/depense', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  data.depenses.push({
    id: Date.now(),
    numero: req.body.numero || '',
    nom: req.body.nom || '',
    type: req.body.type || '',
    phase: req.body.phase || '',
    fournisseur: req.body.fournisseur || '',
    facture: req.body.facture || '',
    date: req.body.date || '',
    description: req.body.description || '',
    quantite: req.body.quantite || '',
    unite: req.body.unite || '',
    montant: req.body.montant || '0'
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/supprimer-depense', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  if (!req.body.id) {
    res.status(400).send('ID manquant');
    return;
  }

  const depense = data.depenses.find(d => String(d.id) === String(req.body.id));

  if (depense && depense.verrouillee === true) {
    res.status(400).send('Suppression interdite pour les lignes automatiques');
    return;
  }

  if (depense && depense.source === 'feuilleHeure') {
    res.status(400).send('Suppression interdite : passe par la feuille d’heure');
    return;
  }

  data.depenses = data.depenses.filter(d => String(d.id) !== String(req.body.id));

  sauvegarder(data);
  res.send('OK');
});

app.post('/modifier-depense', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const depense = data.depenses.find(d => String(d.id) === String(req.body.id));

  if (depense && depense.verrouillee === true) {
    res.status(400).send('Modification interdite pour les lignes automatiques');
    return;
  }

  if (depense && depense.source === 'feuilleHeure') {
    res.status(400).send('Modification interdite : passe par la feuille d’heure');
    return;
  }

  data.depenses = data.depenses.map(d => {
    if (String(d.id) === String(req.body.id)) {
      return {
        ...d,
        numero: req.body.numero || '',
        nom: req.body.nom || '',
        type: req.body.type || '',
        phase: req.body.phase || '',
        fournisseur: req.body.fournisseur || '',
        facture: req.body.facture || '',
        date: req.body.date || '',
        description: req.body.description || '',
        quantite: req.body.quantite || '',
        unite: req.body.unite || '',
        montant: req.body.montant || '0'
      };
    }
    return d;
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/feuille-heure', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const numero = req.body.numero || '';
  const nom = req.body.nom || '';
  const role = req.body.role || 'Technicien';
  const personne = req.body.personne || '';
  const date = req.body.date || '';
  const heures = Number(req.body.heures) || 0;

  const prixHeure = role === 'Cadre'
    ? Number(data.reglages?.prixHeureCadre) || 75
    : Number(data.reglages?.prixHeureTechnicien) || 55;

  const total = heures * prixHeure;
  const id = Date.now();

  data.feuillesHeures.push({
    id,
    numero,
    nom,
    role,
    personne,
    date,
    heures,
    prixHeure,
    total
  });

  data.depenses.push({
    id: Date.now() + 1,
    numero: numero,
    nom: nom,
    type: "Main d'oeuvre",
    phase: 'Initial',
    fournisseur: personne,
    facture: '',
    date: date,
    description: `Feuille d'heure - ${personne} - ${role}`,
    quantite: heures.toString(),
    unite: 'H',
    montant: prixHeure.toString(),
    source: 'feuilleHeure',
    feuilleHeureId: id,
    verrouillee: true
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/modifier-feuille-heure', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const id = req.body.id;
  const role = req.body.role || 'Technicien';

  const prixHeure = role === 'Cadre'
    ? Number(data.reglages?.prixHeureCadre) || 75
    : Number(data.reglages?.prixHeureTechnicien) || 55;

  data.feuillesHeures = data.feuillesHeures.map(fh => {
    if (String(fh.id) === String(id)) {
      const heures = Number(req.body.heures) || 0;
      return {
        ...fh,
        numero: req.body.numero || '',
        nom: req.body.nom || '',
        role: role,
        personne: req.body.personne || '',
        date: req.body.date || '',
        heures: heures,
        prixHeure: prixHeure,
        total: heures * prixHeure
      };
    }
    return fh;
  });

  const feuille = data.feuillesHeures.find(fh => String(fh.id) === String(id));

  if (feuille) {
    const depenseExistante = data.depenses.find(d => String(d.feuilleHeureId) === String(id));

    if (depenseExistante) {
      data.depenses = data.depenses.map(d => {
        if (String(d.feuilleHeureId) === String(id)) {
          return {
            ...d,
            numero: feuille.numero,
            nom: feuille.nom,
            fournisseur: feuille.personne,
            date: feuille.date,
            description: `Feuille d'heure - ${feuille.personne} - ${feuille.role}`,
            quantite: feuille.heures.toString(),
            unite: 'H',
            montant: feuille.prixHeure.toString()
          };
        }
        return d;
      });
    } else {
      data.depenses.push({
        id: Date.now(),
        numero: feuille.numero,
        nom: feuille.nom,
        type: "Main d'oeuvre",
        phase: 'Initial',
        fournisseur: feuille.personne,
        facture: '',
        date: feuille.date,
        description: `Feuille d'heure - ${feuille.personne} - ${feuille.role}`,
        quantite: feuille.heures.toString(),
        unite: 'H',
        montant: feuille.prixHeure.toString(),
        source: 'feuilleHeure',
        feuilleHeureId: feuille.id,
        verrouillee: true
      });
    }
  }

  sauvegarder(data);
  res.send('OK');
});

app.post('/supprimer-feuille-heure', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const id = req.body.id;

  data.feuillesHeures = data.feuillesHeures.filter(fh => String(fh.id) !== String(id));
  data.depenses = data.depenses.filter(d => String(d.feuilleHeureId) !== String(id));

  sauvegarder(data);
  res.send('OK');
});

app.post('/intervention', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  data.interventions.push({
    id: Date.now(),
    numero: req.body.numero || '',
    nom: req.body.nom || '',
    client: req.body.client || '',
    site: req.body.site || '',
    typeInstallation: req.body.typeInstallation || '',
    typeIntervention: req.body.typeIntervention || '',
    datePrevue: req.body.datePrevue || '',
    heurePrevue: req.body.heurePrevue || '',
    technicien: req.body.technicien || '',
    priorite: req.body.priorite || '',
    statut: req.body.statut || '',
    pompe1Heure: req.body.pompe1Heure || '',
    pompe1Kwh: req.body.pompe1Kwh || '',
    pompe1Etat: req.body.pompe1Etat || '',
    pompe2Heure: req.body.pompe2Heure || '',
    pompe2Kwh: req.body.pompe2Kwh || '',
    pompe2Etat: req.body.pompe2Etat || '',
    degrilleur: req.body.degrilleur || '',
    flotteur: req.body.flotteur || '',
    alarme: req.body.alarme || '',
    nettoyage: req.body.nettoyage || '',
    observations: req.body.observations || '',
    anomalies: req.body.anomalies || '',
    signature: req.body.signature || ''
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/modifier-intervention', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  data.interventions = data.interventions.map(i => {
    if (String(i.id) === String(req.body.id)) {
      return {
        ...i,
        numero: req.body.numero || '',
        nom: req.body.nom || '',
        client: req.body.client || '',
        site: req.body.site || '',
        typeInstallation: req.body.typeInstallation || '',
        typeIntervention: req.body.typeIntervention || '',
        datePrevue: req.body.datePrevue || '',
        heurePrevue: req.body.heurePrevue || '',
        technicien: req.body.technicien || '',
        priorite: req.body.priorite || '',
        statut: req.body.statut || '',
        pompe1Heure: req.body.pompe1Heure || '',
        pompe1Kwh: req.body.pompe1Kwh || '',
        pompe1Etat: req.body.pompe1Etat || '',
        pompe2Heure: req.body.pompe2Heure || '',
        pompe2Kwh: req.body.pompe2Kwh || '',
        pompe2Etat: req.body.pompe2Etat || '',
        degrilleur: req.body.degrilleur || '',
        flotteur: req.body.flotteur || '',
        alarme: req.body.alarme || '',
        nettoyage: req.body.nettoyage || '',
        observations: req.body.observations || '',
        anomalies: req.body.anomalies || '',
        signature: req.body.signature || ''
      };
    }
    return i;
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/supprimer-intervention', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  data.interventions = data.interventions.filter(i => String(i.id) !== String(req.body.id));
  sauvegarder(data);
  res.send('OK');
});

app.post('/utilisateur', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  data.utilisateurs.push({
    nom: req.body.nom || '',
    motdepasse: req.body.motdepasse || '',
    role: req.body.role || '',
    droits: {
      chantiers: req.body.chantiers === true,
      depenses: req.body.depenses === true,
      feuillesHeures: req.body.feuillesHeures === true,
      interventions: req.body.interventions === true,
      photos: req.body.photos === true,
      reglages: req.body.reglages === true
    }
  });

  sauvegarder(data);
  res.send('OK');
});

app.post('/restaurer-archive', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const id = req.body.id;

  const archive = (data.archives || []).find(a => String(a.id) === String(id));

  if (!archive) {
    res.status(404).send('Archive introuvable');
    return;
  }

  // remettre les données
  if (archive.chantier) data.chantiers.push(archive.chantier);
  data.depenses.push(...(archive.depenses || []));
  data.feuillesHeures.push(...(archive.feuillesHeures || []));
  data.interventions.push(...(archive.interventions || []));
  data.photos.push(...(archive.photos || []));

  // supprimer archive
  data.archives = data.archives.filter(a => String(a.id) !== String(id));

  sauvegarder(data);
  res.send('OK');
});

app.post('/restaurer-archive', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const id = req.body.id;

  const archive = (data.archives || []).find(a => String(a.id) === String(id));

  if (!archive) {
    res.status(404).send('Archive introuvable');
    return;
  }

  // remettre les données
  if (archive.chantier) data.chantiers.push(archive.chantier);
  data.depenses.push(...(archive.depenses || []));
  data.feuillesHeures.push(...(archive.feuillesHeures || []));
  data.interventions.push(...(archive.interventions || []));
  data.photos.push(...(archive.photos || []));

  // supprimer archive
  data.archives = data.archives.filter(a => String(a.id) !== String(id));

  sauvegarder(data);
  res.send('OK');
});

app.post('/photo', verifierConnexion, upload.single('photo'), (req, res) => {
  const data = lireDonnees();

  data.photos.push({
    chantier: req.body.chantier || '',
    description: req.body.description || '',
    fichier: req.file ? '/uploads/' + req.file.filename : ''
  });

  sauvegarder(data);
  res.redirect('/');
});

app.get('/export-excel', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  let csv = 'NumeroChantier;NomChantier;Type;Phase;Fournisseur;Facture;Date;Description;Quantite;Unite;PrixUnitaire;PrixTotal\n';

  data.depenses.forEach(d => {
    const quantite = Number(d.quantite) || 0;
    const montant = Number(d.montant) || 0;
    const total = quantite * montant;

    csv += `${d.numero || ''};${d.nom || ''};${d.type || ''};${d.phase || ''};${d.fournisseur || ''};${d.facture || ''};${d.date || ''};${d.description || ''};${quantite};${d.unite || ''};${montant};${total}\n`;
  });

  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="depenses_hds.csv"');
  res.send('\uFEFF' + csv);
});

app.get('/export-xlsx', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const rows = data.depenses.map(d => {
    const quantite = Number(d.quantite) || 0;
    const montant = Number(d.montant) || 0;
    const total = quantite * montant;

    return {
      NumeroChantier: d.numero || '',
      NomChantier: d.nom || '',
      Type: d.type || '',
      Phase: d.phase || '',
      Fournisseur: d.fournisseur || '',
      Facture: d.facture || '',
      Date: d.date || '',
      Description: d.description || '',
      Quantite: quantite,
      Unite: d.unite || '',
      PrixUnitaire: montant,
      PrixTotal: total
    };
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Depenses');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', 'attachment; filename="depenses_hds.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  res.send(buffer);
});

app.get('/export-feuilles-heures-xlsx', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const rows = data.feuillesHeures.map(fh => ({
    NumeroChantier: fh.numero || '',
    NomChantier: fh.nom || '',
    Personne: fh.personne || '',
    Role: fh.role || '',
    Date: fh.date || '',
    NombreHeures: Number(fh.heures) || 0,
    PrixHeure: Number(fh.prixHeure) || 0,
    Total: Number(fh.total) || 0
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'FeuillesHeures');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', 'attachment; filename="feuilles_heures_hds.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  res.send(buffer);
});

app.get('/archive/:id', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const archive = (data.archives || []).find(a => String(a.id) === String(req.params.id));
  const totalMarche = Number(archive.chantier?.totalMarcheHT) || 0;

const totalDepenses = (archive.depenses || []).reduce((sum, d) => {
  const q = Number(d.quantite) || 0;
  const m = Number(d.montant) || 0;
  return sum + (q * m);
}, 0);

const reste = totalMarche - totalDepenses;

  if (!archive) {
    res.send('Archive introuvable');
    return;
  }

  res.send(`
    <html>
    <head>
      <meta charset="utf-8">
      <title>Dossier archivé</title>
      <style>
       .menu-mobile {
  display:flex;
  overflow-x:auto;
  background:#1e293b;
  padding:10px;
}

.menu-mobile button {
  flex:1;
  margin:3px;
  background:#2563eb;
  color:white;
  border:none;
  border-radius:6px;
  padding:10px;
  font-size:14px;
}

        body { font-family: Arial; padding:20px; }
        .bloc { border:1px solid #ccc; padding:10px; margin:10px 0; border-radius:8px; }
      </style>
    </head>
    <body>

      <h1>Archive chantier ${archive.numero}</h1>

<div class="bloc">
  <strong>Montant marché HT :</strong> ${totalMarche} €<br>
  <strong>Total dépenses :</strong> ${totalDepenses} €<br>
  <strong>Reste :</strong> ${reste} €
</div>

      <div class="bloc">
        <h3>Dépenses</h3>
        ${(archive.depenses || []).map(d => `
          <div>${d.description || ''} - ${d.montant || 0}€</div>
        `).join('')}
      </div>

      <div class="bloc">
        <h3>Feuilles d'heure</h3>
        ${(archive.feuillesHeures || []).map(f => `
          <div>${f.personne || ''} - ${f.heures || 0}h</div>
        `).join('')}
      </div>

      <div class="bloc">
        <h3>Interventions</h3>
        ${(archive.interventions || []).map(i => `
          <div>${i.typeIntervention || ''} - ${i.datePrevue || ''}</div>
        `).join('')}
      </div>

      <div class="bloc">
        <h3>Photos</h3>
        ${(archive.photos || []).map(p => `
          <img src="${p.fichier}" style="max-width:200px;margin:5px;">
        `).join('')}
      </div>

    </body>
    </html>
  `);
});

app.get('/export-interventions-xlsx', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const rows = data.interventions.map(i => ({
    NumeroChantier: i.numero || '',
    NomChantier: i.nom || '',
    Client: i.client || '',
    Site: i.site || '',
    TypeInstallation: i.typeInstallation || '',
    TypeIntervention: i.typeIntervention || '',
    DatePrevue: i.datePrevue || '',
    HeurePrevue: i.heurePrevue || '',
    Technicien: i.technicien || '',
    Priorite: i.priorite || '',
    Statut: i.statut || '',
    Pompe1Heure: i.pompe1Heure || '',
    Pompe1Kwh: i.pompe1Kwh || '',
    Pompe1Etat: i.pompe1Etat || '',
    Pompe2Heure: i.pompe2Heure || '',
    Pompe2Kwh: i.pompe2Kwh || '',
    Pompe2Etat: i.pompe2Etat || '',
    Degrilleur: i.degrilleur || '',
    Flotteur: i.flotteur || '',
    Alarme: i.alarme || '',
    Nettoyage: i.nettoyage || '',
    Observations: i.observations || '',
    Anomalies: i.anomalies || ''
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Interventions');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', 'attachment; filename="interventions_hds.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  res.send(buffer);
});

app.get('/export-chantier-xlsx/:numero', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const numero = req.params.numero;

  const chantier = data.chantiers.find(c => String(c.numero) === String(numero));

  if (!chantier) {
    res.status(404).send('Chantier introuvable');
    return;
  }

  const depenses = data.depenses.filter(d => String(d.numero) === String(numero));

  const rows = depenses.map(d => {
    const quantite = Number(d.quantite) || 0;
    const montant = Number(d.montant) || 0;
    const total = quantite * montant;

    return {
      NumeroChantier: d.numero || '',
      NomChantier: d.nom || '',
      Type: d.type || '',
      Phase: d.phase || '',
      Fournisseur: d.fournisseur || '',
      Facture: d.facture || '',
      Date: d.date || '',
      Description: d.description || '',
      Quantite: quantite,
      Unite: d.unite || '',
      PrixUnitaire: montant,
      PrixTotal: total
    };
  });

  const totalDepenses = depenses.reduce((somme, d) => {
    const quantite = Number(d.quantite) || 0;
    const montant = Number(d.montant) || 0;
    return somme + (quantite * montant);
  }, 0);

  const prixMarcheHT = Number(chantier.totalMarcheHT) || 0;
  const reste = prixMarcheHT - totalDepenses;

  const resume = [
    {
      NumeroChantier: chantier.numero || '',
      NomChantier: chantier.nom || '',
      DateChantier: chantier.date || '',
      MaitreOuvrage: chantier.maitreOuvrage || '',
      MaitreOeuvre: chantier.maitreOeuvre || '',
      Lieu: chantier.lieu || '',
      PrixMarcheHT: prixMarcheHT,
      TotalDepenses: totalDepenses,
      Reste: reste
    }
  ];

  const wb = XLSX.utils.book_new();

  const wsResume = XLSX.utils.json_to_sheet(resume);
  XLSX.utils.book_append_sheet(wb, wsResume, 'Resume');

  const wsDepenses = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, wsDepenses, 'Depenses');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', `attachment; filename="chantier_${numero}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  res.send(buffer);
});

app.get('/rapport-intervention/:id', verifierConnexion, (req, res) => {
  const data = lireDonnees();
  const id = req.params.id;

  const i = data.interventions.find(x => String(x.id) === String(id));

  if (!i) {
    res.send('Intervention introuvable');
    return;
  }

  res.send(`
    <html>
    <head>
      <meta charset="utf-8">
      <title>Rapport intervention</title>
      <style>
        body { font-family: Arial; padding:20px; }
        h1 { text-align:center; }
        .bloc { border:1px solid #ccc; padding:10px; margin:10px 0; border-radius:8px; }
        .signature { margin-top:20px; }
        img { max-width:300px; border:1px solid #000; }
        button { padding:10px; margin-top:20px; }
      </style>
    </head>
    <body>
      <h1>RAPPORT D’INTERVENTION</h1>

      <div class="bloc">
        <strong>Chantier :</strong> ${i.numero} - ${i.nom}<br>
        <strong>Client :</strong> ${i.client}<br>
        <strong>Site :</strong> ${i.site}<br>
        <strong>Date :</strong> ${i.datePrevue} ${i.heurePrevue}<br>
        <strong>Technicien :</strong> ${i.technicien}<br>
      </div>

      <div class="bloc">
        <strong>Type installation :</strong> ${i.typeInstallation}<br>
        <strong>Type intervention :</strong> ${i.typeIntervention}<br>
      </div>

      <div class="bloc">
        <h3>POMPE 1</h3>
        Heures : ${i.pompe1Heure}<br>
        kWh : ${i.pompe1Kwh}<br>
        Etat : ${i.pompe1Etat}
      </div>

      <div class="bloc">
        <h3>POMPE 2</h3>
        Heures : ${i.pompe2Heure}<br>
        kWh : ${i.pompe2Kwh}<br>
        Etat : ${i.pompe2Etat}
      </div>

      <div class="bloc">
        Dégrilleur : ${i.degrilleur}<br>
        Flotteur : ${i.flotteur}<br>
        Alarme : ${i.alarme}<br>
        Nettoyage : ${i.nettoyage}
      </div>

      <div class="bloc">
        <strong>Observations :</strong><br>
        ${i.observations}
      </div>

      <div class="bloc">
        <strong>Anomalies :</strong><br>
        ${i.anomalies}
      </div>

      <div class="bloc signature">
        <strong>Signature client :</strong><br>
        ${i.signature ? `<img src="${i.signature}" />` : 'Non signée'}
      </div>

      <button onclick="window.print()">🖨️ Imprimer / PDF</button>
    </body>
    </html>
  `);
});

app.get('/export-word-intervention/:id', verifierConnexion, async (req, res) => {
  const data = lireDonnees();
  const id = req.params.id;

  const i = data.interventions.find(x => String(x.id) === String(id));

  if (!i) {
    res.status(404).send('Intervention introuvable');
    return;
  }

  const enfants = [];

  enfants.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "RAPPORT D’INTERVENTION HDS",
          bold: true,
          size: 32
        })
      ]
    })
  );

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Chantier : ", bold: true }), new TextRun(`${i.numero || ''} - ${i.nom || ''}`)]
  }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Client : ", bold: true }), new TextRun(`${i.client || ''}`)]
  }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Site : ", bold: true }), new TextRun(`${i.site || ''}`)]
  }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Date : ", bold: true }), new TextRun(`${i.datePrevue || ''} ${i.heurePrevue || ''}`)]
  }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Technicien : ", bold: true }), new TextRun(`${i.technicien || ''}`)]
  }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Type installation : ", bold: true }), new TextRun(`${i.typeInstallation || ''}`)]
  }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "Type intervention : ", bold: true }), new TextRun(`${i.typeIntervention || ''}`)]
  }));

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "POMPE 1", bold: true, size: 26 })]
  }));

  enfants.push(new Paragraph({ text: `Heures : ${i.pompe1Heure || ''}` }));
  enfants.push(new Paragraph({ text: `kWh : ${i.pompe1Kwh || ''}` }));
  enfants.push(new Paragraph({ text: `État : ${i.pompe1Etat || ''}` }));

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "POMPE 2", bold: true, size: 26 })]
  }));

  enfants.push(new Paragraph({ text: `Heures : ${i.pompe2Heure || ''}` }));
  enfants.push(new Paragraph({ text: `kWh : ${i.pompe2Kwh || ''}` }));
  enfants.push(new Paragraph({ text: `État : ${i.pompe2Etat || ''}` }));

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "CONTRÔLES", bold: true, size: 26 })]
  }));

  enfants.push(new Paragraph({ text: `Panier dégrilleur : ${i.degrilleur || ''}` }));
  enfants.push(new Paragraph({ text: `Flotteurs testés : ${i.flotteur || ''}` }));
  enfants.push(new Paragraph({ text: `Alarme testée : ${i.alarme || ''}` }));
  enfants.push(new Paragraph({ text: `Nettoyage : ${i.nettoyage || ''}` }));

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "OBSERVATIONS", bold: true, size: 26 })]
  }));

  enfants.push(new Paragraph({ text: `${i.observations || ''}` }));

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "ANOMALIES", bold: true, size: 26 })]
  }));

  enfants.push(new Paragraph({ text: `${i.anomalies || ''}` }));

  enfants.push(new Paragraph({ text: "" }));

  enfants.push(new Paragraph({
    children: [new TextRun({ text: "SIGNATURE CLIENT", bold: true, size: 26 })]
  }));

  if (i.signature && i.signature.startsWith('data:image')) {
    try {
      const base64Data = i.signature.replace(/^data:image\/png;base64,/, '');
      const imageBuffer = Buffer.from(base64Data, 'base64');

      enfants.push(
        new Paragraph({
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: {
                width: 200,
                height: 100
              }
            })
          ]
        })
      );
    } catch (e) {
      enfants.push(new Paragraph({ text: "Signature présente mais impossible à afficher." }));
    }
  } else {
    enfants.push(new Paragraph({ text: "Aucune signature" }));
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: enfants
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);

  res.setHeader('Content-Disposition', `attachment; filename="rapport_intervention_${i.numero || 'HDS'}.docx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

  res.send(buffer);
});

app.get('/planning-pdf', verifierConnexion, (req, res) => {
  const data = lireDonnees();

  const interventionsTriees = [...data.interventions].sort((a, b) => {
    const dateA = new Date((a.datePrevue || '') + ' ' + (a.heurePrevue || '00:00'));
    const dateB = new Date((b.datePrevue || '') + ' ' + (b.heurePrevue || '00:00'));
    return dateA - dateB;
  });

  const planning = {};
  interventionsTriees.forEach(i => {
    const date = i.datePrevue || 'Sans date';
    if (!planning[date]) planning[date] = [];
    planning[date].push(i);
  });

  const htmlPlanningPdf = Object.keys(planning).map(date => {
    const interventionsJour = planning[date].map(i => `
      <tr>
        <td>${i.heurePrevue || ''}</td>
        <td>${i.numero || ''}</td>
        <td>${i.nom || ''}</td>
        <td>${i.technicien || ''}</td>
        <td>${i.typeInstallation || ''}</td>
        <td>${i.typeIntervention || ''}</td>
        <td>${i.statut || ''}</td>
        <td>${i.priorite || ''}</td>
      </tr>
    `).join('');

    return `
      <h2 style="margin-top:30px;">📅 ${date}</h2>
      <table style="width:100%; border-collapse:collapse; margin-bottom:20px;">
        <thead>
          <tr style="background:#e5e7eb;">
            <th style="border:1px solid #999;padding:8px;">Heure</th>
            <th style="border:1px solid #999;padding:8px;">Numéro</th>
            <th style="border:1px solid #999;padding:8px;">Chantier</th>
            <th style="border:1px solid #999;padding:8px;">Technicien</th>
            <th style="border:1px solid #999;padding:8px;">Installation</th>
            <th style="border:1px solid #999;padding:8px;">Intervention</th>
            <th style="border:1px solid #999;padding:8px;">Statut</th>
            <th style="border:1px solid #999;padding:8px;">Priorité</th>
          </tr>
        </thead>
        <tbody>
          ${interventionsJour}
        </tbody>
      </table>
    `;
  }).join('') || '<p>Aucune intervention</p>';

  res.send(`
    <html>
    <head>
      <meta charset="utf-8">
      <title>Planning HDS</title>
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; color: #111827; }
        h1 { text-align: center; margin-bottom: 30px; }
        table { font-size: 14px; }
        th, td { border: 1px solid #999; padding: 8px; text-align: left; }
        .topbar { margin-bottom: 20px; text-align: center; }
        .btn {
          padding: 10px 16px;
          border: none;
          border-radius: 6px;
          cursor: pointer;
          background: #2563eb;
          color: white;
          margin: 0 5px;
        }
        @media print {
          .topbar { display: none; }
        }
      </style>
    </head>
    <body>
      <div class="topbar">
        <button class="btn" onclick="window.print()">Imprimer / Enregistrer en PDF</button>
        <button class="btn" onclick="window.close()">Fermer</button>
      </div>

      <h1>PLANNING DES INTERVENTIONS HDS</h1>

      ${htmlPlanningPdf}
    </body>
    </html>
  `);
});

app.listen(PORT, '0.0.0.0', () => {
  console.log('HDS Pro lancé sur port ' + PORT);
});