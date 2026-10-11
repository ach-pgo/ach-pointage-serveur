// ACH ProGestions & Outils — Serveur de pointage et timesheet
const express = require("express");
const Database = require("better-sqlite3");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const https = require("https");

const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const FUSEAU = process.env.FUSEAU || "Africa/Lagos"; // Bénin = UTC+1
fs.mkdirSync(DATA_DIR, { recursive: true });

const TURSO_URL = process.env.TURSO_DATABASE_URL || "https://ach-pointage-db-ach-pgo.aws-eu-west-1.turso.io";
const TURSO_TOKEN = process.env.TURSO_AUTH_TOKEN || "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3OTE3MjgzODUsImlkIjoiMDFhMTJiNGItOTcwMS03ZmI3LWFkN2EtNzEzZTQ3NGM5NDZlIiwia2lkIjoiRElBa3RvTmpxQjRyckY5RS1nUWRrU2l2bloxZnJReTU4a2gyVkZ2aXQ3TSIsInJpZCI6IjViYjJmNDA5LTc0YTktNGUyYS1iMWJhLTFiZDk3YWE2OTJkNSJ9.z1fKesSsyHp4IXaYptWelyZKKsiE4p6N6-CBlIqyCqJN02fUN5ZtbLAO9FMcoltGFbdjVR4sEHL9wbIYhNZbAg";

function queryTurso(sql, args = []) {
  return new Promise(resolve => {
    try {
      const endpoint = TURSO_URL.replace(/^libsql:\/\//, "https://").replace(/\/$/, "") + "/v2/pipeline";
      const payload = JSON.stringify({ requests: [{ type: "execute", stmt: { sql, args } }] });
      const req = https.request(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": "Bearer " + TURSO_TOKEN }
      }, res => {
        let raw = "";
        res.on("data", c => raw += c);
        res.on("end", () => {
          try {
            const j = JSON.parse(raw);
            resolve(j?.results?.[0]?.response?.result || null);
          } catch { resolve(null); }
        });
      });
      req.on("error", () => resolve(null));
      req.write(payload);
      req.end();
    } catch { resolve(null); }
  });
}

const dbPath = path.join(DATA_DIR, "ach.db");

// Restauration synchrone avant initialisation de la base si disponible sur Turso
try {
  const child = require("child_process");
  const restoreScript = `
    const https = require('https'); const fs = require('fs');
    const u = ('${TURSO_URL}'.replace(/^libsql:\\/\\//, 'https://').replace(/\\/$/, '')) + '/v2/pipeline';
    const postData = JSON.stringify({ requests: [{ type: 'execute', stmt: { sql: 'SELECT donnees_base64 FROM _ach_cloud_backup WHERE id=1' } }] });
    const req = https.request(u, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ${TURSO_TOKEN}' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => {
        try {
          const j = JSON.parse(d);
          const b64 = j?.results?.[0]?.response?.result?.rows?.[0]?.[0]?.value;
          if (b64 && b64.length > 50) {
            fs.writeFileSync('${dbPath.replace(/\\/g, "/")}', Buffer.from(b64, 'base64'));
            console.log('Restauré depuis Turso avec succès (' + b64.length + ' octets base64)');
          }
        } catch(e) {}
      });
    });
    req.on('error', () => {});
    req.write(postData); req.end();
  `;
  child.execSync(`node -e "${restoreScript.replace(/\n/g, ' ')}"`, { timeout: 8000 });
} catch(e) {}

const db = new Database(dbPath);
db.pragma("journal_mode = WAL");

let sauvTimeout = null;
function synchroniserTurso() {
  if (sauvTimeout) clearTimeout(sauvTimeout);
  sauvTimeout = setTimeout(async () => {
    try {
      const b64 = db.serialize().toString("base64");
      const n = new Date().toISOString().slice(0, 19).replace("T", " ");
      await queryTurso(`
        CREATE TABLE IF NOT EXISTS _ach_cloud_backup (id INTEGER PRIMARY KEY, date_maj TEXT, donnees_base64 TEXT)
      `);
      await queryTurso(`
        INSERT INTO _ach_cloud_backup(id, date_maj, donnees_base64) VALUES(1, ?, ?)
        ON CONFLICT(id) DO UPDATE SET date_maj=excluded.date_maj, donnees_base64=excluded.donnees_base64
      `, [{ type: "text", value: n }, { type: "text", value: b64 }]);
      console.log("Sauvegarde permanente effectuée sur Turso Cloud (" + n + ")");
    } catch (e) {
      console.warn("Avertissement sauvegarde Turso:", e.message);
    }
  }, 1000);
}

// Sauvegarde automatique toutes les 15 minutes
setInterval(synchroniserTurso, 15 * 60 * 1000);

db.exec(`
CREATE TABLE IF NOT EXISTS utilisateurs (
  id INTEGER PRIMARY KEY AUTOINCREMENT, nom TEXT NOT NULL, identifiant TEXT NOT NULL UNIQUE,
  poste TEXT DEFAULT '', mdp_hash TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'employe',
  actif INTEGER NOT NULL DEFAULT 1, doit_changer_mdp INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS connexions (
  id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, date TEXT NOT NULL, heure TEXT NOT NULL,
  type TEXT NOT NULL, pc TEXT DEFAULT '', heure_pc TEXT DEFAULT '', ts INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS idx_conn ON connexions(user_id, date);
CREATE TABLE IF NOT EXISTS saisies (
  id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, date TEXT NOT NULL, activite TEXT DEFAULT '',
  client TEXT DEFAULT '', tache TEXT DEFAULT '', description TEXT DEFAULT '', heures REAL NOT NULL DEFAULT 0);
CREATE INDEX IF NOT EXISTS idx_saisies ON saisies(user_id, date);
CREATE TABLE IF NOT EXISTS validations (
  user_id INTEGER NOT NULL, mois TEXT NOT NULL, statut TEXT NOT NULL, valide_par TEXT DEFAULT '', date_validation TEXT DEFAULT '',
  PRIMARY KEY (user_id, mois));
CREATE TABLE IF NOT EXISTS parametres (cle TEXT PRIMARY KEY, valeur TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS activites_pc (
  id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, date TEXT NOT NULL,
  pc TEXT DEFAULT '', heure_debut TEXT DEFAULT '', heure_dernier_signal TEXT DEFAULT '',
  derniere_action TEXT DEFAULT '', inactif_secondes INTEGER DEFAULT 0,
  arret_detecte INTEGER DEFAULT 0, heure_arret TEXT DEFAULT '', ts_maj INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS idx_actpc ON activites_pc(user_id, date);
`);

const getParam = (k, d) => { const r = db.prepare("SELECT valeur FROM parametres WHERE cle=?").get(k); return r ? r.valeur : d; };
const setParam = (k, v) => {
  db.prepare("INSERT INTO parametres(cle,valeur) VALUES(?,?) ON CONFLICT(cle) DO UPDATE SET valeur=excluded.valeur").run(k, v);
  synchroniserTurso();
};
let SECRET = process.env.JWT_SECRET || getParam("jwt_secret", null);
if (!SECRET) { SECRET = crypto.randomBytes(48).toString("hex"); setParam("jwt_secret", SECRET); }
if (!getParam("activites", null)) setParam("activites", JSON.stringify(["Deméya Service Conseil", "EHPAD Invest", "Autre"]));
if (!getParam("heure_attendue", null)) setParam("heure_attendue", "08:00");

if (db.prepare("SELECT COUNT(*) c FROM utilisateurs").get().c === 0) {
  const u = process.env.ADMIN_USER || "admin", p = process.env.ADMIN_PASSWORD || "ChangeMoi123!";
  db.prepare("INSERT INTO utilisateurs(nom,identifiant,poste,mdp_hash,role,doit_changer_mdp) VALUES(?,?,?,?,?,1)")
    .run(process.env.ADMIN_NOM || "Hosiane AYABA", u, "Direction", bcrypt.hashSync(p, 10), "admin");
  console.log(`Compte administrateur créé : identifiant "${u}" — changez le mot de passe dès la première connexion.`);
  synchroniserTurso();
}

// ---------- utilitaires
const MOIS = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
function maintenantLocal() {
  const p = Object.fromEntries(new Intl.DateTimeFormat("en-CA", { timeZone: FUSEAU, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit" }).formatToParts(new Date()).map(x => [x.type, x.value]));
  return { date: `${p.year}-${p.month}-${p.day}`, heure: `${p.hour}:${p.minute}:${p.second}` };
}
const moisValide = m => /^\d{4}-(0[1-9]|1[0-2])$/.test(m || "");
const dateValide = d => /^\d{4}-\d{2}-\d{2}$/.test(d || "") && !isNaN(Date.parse(d));
const borne = m => [`${m}-01`, `${m}-31`];
const hhmm = s => (s || "").slice(0, 5);

const app = express();
app.use(express.json({ limit: "200kb" }));
app.use(express.static(path.join(__dirname, "public")));

// ---------- authentification
const echecs = new Map();
app.post("/api/login", (req, res) => {
  const { identifiant = "", mdp = "" } = req.body || {};
  const cle = req.ip + "|" + identifiant.toLowerCase(), e = echecs.get(cle);
  if (e && e.n >= 6 && Date.now() - e.t < 15 * 60000) return res.status(429).json({ erreur: "Trop d'essais. Réessayez dans 15 minutes." });
  const u = db.prepare("SELECT * FROM utilisateurs WHERE lower(identifiant)=lower(?)").get(identifiant.trim());
  if (!u || !u.actif || !bcrypt.compareSync(mdp, u.mdp_hash)) {
    echecs.set(cle, { n: (e && Date.now() - e.t < 15 * 60000 ? e.n : 0) + 1, t: Date.now() });
    return res.status(401).json({ erreur: "Identifiant ou mot de passe incorrect." });
  }
  echecs.delete(cle);
  res.json({ token: jwt.sign({ id: u.id }, SECRET, { expiresIn: "60d" }) });
});
function auth(req, res, next) {
  const t = (req.headers.authorization || "").replace("Bearer ", "");
  try {
    const { id } = jwt.verify(t, SECRET);
    const u = db.prepare("SELECT * FROM utilisateurs WHERE id=? AND actif=1").get(id);
    if (!u) throw 0; req.user = u; next();
  } catch { res.status(401).json({ erreur: "Session expirée." }); }
}
const admin = (req, res, next) => req.user.role === "admin" ? next() : res.status(403).json({ erreur: "Réservé à l'administratrice." });
// cible : l'employé lui-même, ou n'importe qui pour l'admin
function cible(req) {
  const id = Number(req.query.user_id || req.body?.user_id || req.user.id);
  if (id !== req.user.id && req.user.role !== "admin") return null;
  return db.prepare("SELECT * FROM utilisateurs WHERE id=?").get(id);
}
const publicUser = u => ({ id: u.id, nom: u.nom, identifiant: u.identifiant, poste: u.poste, role: u.role, actif: !!u.actif, doit_changer_mdp: !!u.doit_changer_mdp });

app.get("/api/me", auth, (req, res) => res.json({ ...publicUser(req.user), ...maintenantLocal(), activites: JSON.parse(getParam("activites")), heure_attendue: getParam("heure_attendue") }));
app.post("/api/mon-mot-de-passe", auth, (req, res) => {
  const { ancien = "", nouveau = "" } = req.body || {};
  if (!bcrypt.compareSync(ancien, req.user.mdp_hash)) return res.status(400).json({ erreur: "Ancien mot de passe incorrect." });
  if (nouveau.length < 8) return res.status(400).json({ erreur: "Le nouveau mot de passe doit contenir au moins 8 caractères." });
  db.prepare("UPDATE utilisateurs SET mdp_hash=?, doit_changer_mdp=0 WHERE id=?").run(bcrypt.hashSync(nouveau, 10), req.user.id);
  synchroniserTurso();
  res.json({ ok: true });
});

// ---------- pointage
app.post("/api/connexion", auth, (req, res) => {
  const type = req.body?.type === "manuel" ? "manuel" : "auto";
  const n = maintenantLocal(), ts = Date.now();
  if (type === "auto") {
    const dernier = db.prepare("SELECT ts FROM connexions WHERE user_id=? AND type='auto' ORDER BY ts DESC LIMIT 1").get(req.user.id);
    if (dernier && ts - dernier.ts < 20 * 60000) return res.json({ ok: true, doublon: true, arrivee: arrivee(req.user.id, n.date) });
  }
  db.prepare("INSERT INTO connexions(user_id,date,heure,type,pc,heure_pc,ts) VALUES(?,?,?,?,?,?,?)")
    .run(req.user.id, n.date, n.heure, type, String(req.body?.pc || "").slice(0, 80), String(req.body?.heure_pc || "").slice(0, 40), ts);
  synchroniserTurso();
  res.json({ ok: true, arrivee: arrivee(req.user.id, n.date) });
});
const arrivee = (uid, date) => hhmm(db.prepare("SELECT MIN(heure) h FROM connexions WHERE user_id=? AND date=?").get(uid, date).h);
app.get("/api/aujourdhui", auth, (req, res) => {
  const n = maintenantLocal();
  res.json({ date: n.date, arrivee: arrivee(req.user.id, n.date), connexions: db.prepare("SELECT heure,type,pc FROM connexions WHERE user_id=? AND date=? ORDER BY heure").all(req.user.id, n.date) });
});
app.get("/api/connexions", auth, (req, res) => {
  const u = cible(req); if (!u) return res.status(403).json({ erreur: "Accès refusé." });
  const mois = req.query.mois;
  if (!moisValide(mois)) return res.status(400).json({ erreur: "Mois invalide." });
  const [a, b] = borne(mois);
  const rows = db.prepare("SELECT c.date,c.heure,c.type,c.pc,c.heure_pc,u.nom FROM connexions c JOIN utilisateurs u ON u.id=c.user_id WHERE (?=0 OR c.user_id=?) AND c.date BETWEEN ? AND ? ORDER BY c.date DESC,c.heure DESC LIMIT 2000")
    .all(req.user.role === "admin" && !req.query.user_id ? 0 : u.id, u.id, a, b);
  res.json(rows);
});

// ---------- synchronisation différée (mode hors-ligne)
app.post("/api/sync-offline", auth, (req, res) => {
  const { connexions = [], activites = [] } = req.body || {};
  let insConn = 0, insAct = 0;

  for (const c of connexions) {
    if (!c.date || !c.heure) continue;
    const ts = c.ts || Date.now();
    const exist = db.prepare("SELECT id FROM connexions WHERE user_id=? AND date=? AND heure=? AND type=?").get(req.user.id, c.date, c.heure, c.type || "auto");
    if (!exist) {
      db.prepare("INSERT INTO connexions(user_id,date,heure,type,pc,heure_pc,ts) VALUES(?,?,?,?,?,?,?)")
        .run(req.user.id, c.date, c.heure, c.type || "auto", String(c.pc || "").slice(0, 80), String(c.heure_pc || c.heure).slice(0, 40), ts);
      insConn++;
    }
  }

  for (const a of activites) {
    if (!a.date) continue;
    const exist = db.prepare("SELECT * FROM activites_pc WHERE user_id=? AND date=?").get(req.user.id, a.date);
    const ts = a.ts || Date.now();
    const inactif = Math.max(0, parseInt(a.inactif_secondes, 10) || 0);
    const arret = a.arret ? 1 : 0;
    const heureDebut = a.heure_debut || a.heure || "";
    const heureAction = a.derniere_action || a.heure_dernier_signal || heureDebut;
    const heureSignal = a.heure_dernier_signal || a.heure || heureDebut;
    const heureArr = a.heure_arret || (arret ? heureSignal : "");

    if (!exist) {
      db.prepare(`
        INSERT INTO activites_pc(user_id, date, pc, heure_debut, heure_dernier_signal, derniere_action, inactif_secondes, arret_detecte, heure_arret, ts_maj)
        VALUES(?,?,?,?,?,?,?,?,?,?)
      `).run(req.user.id, a.date, String(a.pc || "").slice(0, 80), heureDebut, heureSignal, heureAction, inactif, arret, heureArr, ts);
      insAct++;
    } else {
      const hDeb = exist.heure_debut && exist.heure_debut < heureDebut ? exist.heure_debut : (heureDebut || exist.heure_debut);
      const hSig = exist.heure_dernier_signal && exist.heure_dernier_signal > heureSignal ? exist.heure_dernier_signal : heureSignal;
      const hAct = exist.derniere_action && exist.derniere_action > heureAction ? exist.derniere_action : heureAction;
      const maxInactif = Math.max(exist.inactif_secondes || 0, inactif);
      db.prepare(`
        UPDATE activites_pc SET
          pc=?, heure_debut=?, heure_dernier_signal=?, derniere_action=?, inactif_secondes=?,
          arret_detecte=?, heure_arret=?, ts_maj=?
        WHERE id=?
      `).run(a.pc || exist.pc, hDeb, hSig, hAct, maxInactif, arret || exist.arret_detecte, heureArr || exist.heure_arret, ts, exist.id);
      insAct++;
    }
  }

  if (insConn > 0 || insAct > 0) synchroniserTurso();
  res.json({ ok: true, synchronises: { connexions: insConn, activites: insAct } });
});

// ---------- suivi activité & arrêt PC
app.post("/api/activite-pc", auth, (req, res) => {
  const n = maintenantLocal(), ts = Date.now();
  const pc = String(req.body?.pc || "").slice(0, 80);
  const inactif = Math.max(0, parseInt(req.body?.inactif_secondes, 10) || 0);
  const arret = req.body?.arret ? 1 : 0;
  const exist = db.prepare("SELECT * FROM activites_pc WHERE user_id=? AND date=?").get(req.user.id, n.date);

  // Déduire l'heure de la dernière action (clavier/souris)
  const tsDerniereAction = ts - (inactif * 1000);
  const dAction = new Date(tsDerniereAction);
  const heureAction = `${String(dAction.getHours()).padStart(2, '0')}:${String(dAction.getMinutes()).padStart(2, '0')}:${String(dAction.getSeconds()).padStart(2, '0')}`;

  if (!exist) {
    db.prepare(`
      INSERT INTO activites_pc(user_id, date, pc, heure_debut, heure_dernier_signal, derniere_action, inactif_secondes, arret_detecte, heure_arret, ts_maj)
      VALUES(?,?,?,?,?,?,?,?,?,?)
    `).run(req.user.id, n.date, pc, n.heure, n.heure, heureAction, inactif, arret, arret ? n.heure : "", ts);
  } else {
    const heureDebut = exist.heure_debut || n.heure;
    const heureArr = arret ? n.heure : exist.heure_arret;
    const maxInactif = Math.max(exist.inactif_secondes, inactif);
    db.prepare(`
      UPDATE activites_pc SET
        pc=?, heure_dernier_signal=?, derniere_action=?, inactif_secondes=?,
        arret_detecte=?, heure_arret=?, ts_maj=?
      WHERE id=?
    `).run(pc || exist.pc, n.heure, heureAction, maxInactif, arret || exist.arret_detecte, heureArr, ts, exist.id);
  }
  synchroniserTurso();
  res.json({ ok: true });
});

// ---------- saisies (timesheet)
const statutMois = (uid, mois) => (db.prepare("SELECT * FROM validations WHERE user_id=? AND mois=?").get(uid, mois) || { statut: "Non validé" });
function saisiesMois(uid, mois) {
  const [a, b] = borne(mois);
  return db.prepare("SELECT s.*, (SELECT MIN(heure) FROM connexions c WHERE c.user_id=s.user_id AND c.date=s.date) arrivee FROM saisies s WHERE s.user_id=? AND s.date BETWEEN ? AND ? ORDER BY s.date,s.id").all(uid, a, b)
    .map(r => ({ ...r, arrivee: hhmm(r.arrivee) }));
}
app.get("/api/saisies", auth, (req, res) => {
  const u = cible(req); if (!u) return res.status(403).json({ erreur: "Accès refusé." });
  if (!moisValide(req.query.mois)) return res.status(400).json({ erreur: "Mois invalide." });
  res.json({ utilisateur: publicUser(u), validation: statutMois(u.id, req.query.mois), saisies: saisiesMois(u.id, req.query.mois) });
});
function lireSaisie(b) {
  const h = Number(String(b.heures).replace(",", "."));
  if (!dateValide(b.date)) return { erreur: "Date invalide." };
  if (!(h >= 0 && h <= 24)) return { erreur: "Heures invalides (0 à 24)." };
  const t = (x, n) => String(x || "").trim().slice(0, n);
  return { v: [b.date, t(b.activite, 80), t(b.client, 120), t(b.tache, 160), t(b.description, 1000), h] };
}
function verrou(req, uid, date) {
  return req.user.role !== "admin" && statutMois(uid, date.slice(0, 7)).statut === "Validé";
}
app.post("/api/saisies", auth, (req, res) => {
  const u = cible(req); if (!u) return res.status(403).json({ erreur: "Accès refusé." });
  const s = lireSaisie(req.body || {}); if (s.erreur) return res.status(400).json(s);
  if (verrou(req, u.id, s.v[0])) return res.status(423).json({ erreur: "Ce mois est validé : modification impossible." });
  const r = db.prepare("INSERT INTO saisies(user_id,date,activite,client,tache,description,heures) VALUES(?,?,?,?,?,?,?)").run(u.id, ...s.v);
  synchroniserTurso();
  res.json({ id: r.lastInsertRowid });
});
function saisieAutorisee(req, res) {
  const s = db.prepare("SELECT * FROM saisies WHERE id=?").get(req.params.id);
  if (!s || (s.user_id !== req.user.id && req.user.role !== "admin")) { res.status(404).json({ erreur: "Introuvable." }); return null; }
  if (verrou(req, s.user_id, s.date)) { res.status(423).json({ erreur: "Ce mois est validé : modification impossible." }); return null; }
  return s;
}
app.put("/api/saisies/:id", auth, (req, res) => {
  const s = saisieAutorisee(req, res); if (!s) return;
  const n = lireSaisie(req.body || {}); if (n.erreur) return res.status(400).json(n);
  if (verrou(req, s.user_id, n.v[0])) return res.status(423).json({ erreur: "Le mois choisi est validé." });
  db.prepare("UPDATE saisies SET date=?,activite=?,client=?,tache=?,description=?,heures=? WHERE id=?").run(...n.v, s.id);
  synchroniserTurso();
  res.json({ ok: true });
});
app.delete("/api/saisies/:id", auth, (req, res) => {
  const s = saisieAutorisee(req, res); if (!s) return;
  db.prepare("DELETE FROM saisies WHERE id=?").run(s.id);
  synchroniserTurso();
  res.json({ ok: true });
});

// ---------- administration
app.post("/api/validation", auth, admin, (req, res) => {
  const { user_id, mois, statut } = req.body || {};
  if (!moisValide(mois) || !["Validé", "Non validé"].includes(statut)) return res.status(400).json({ erreur: "Données invalides." });
  db.prepare("INSERT INTO validations(user_id,mois,statut,valide_par,date_validation) VALUES(?,?,?,?,?) ON CONFLICT(user_id,mois) DO UPDATE SET statut=excluded.statut,valide_par=excluded.valide_par,date_validation=excluded.date_validation")
    .run(user_id, mois, statut, req.user.nom, maintenantLocal().date);
  synchroniserTurso();
  res.json({ ok: true });
});
app.get("/api/utilisateurs", auth, admin, (req, res) => res.json(db.prepare("SELECT * FROM utilisateurs ORDER BY role,nom").all().map(publicUser)));
app.post("/api/utilisateurs", auth, admin, (req, res) => {
  const { nom = "", identifiant = "", poste = "", mdp = "", role = "employe" } = req.body || {};
  if (!nom.trim() || !identifiant.trim()) return res.status(400).json({ erreur: "Nom et identifiant obligatoires." });
  if (mdp.length < 8) return res.status(400).json({ erreur: "Mot de passe provisoire : 8 caractères minimum." });
  try {
    const r = db.prepare("INSERT INTO utilisateurs(nom,identifiant,poste,mdp_hash,role,doit_changer_mdp) VALUES(?,?,?,?,?,1)")
      .run(nom.trim(), identifiant.trim(), poste.trim(), bcrypt.hashSync(mdp, 10), role === "admin" ? "admin" : "employe");
    synchroniserTurso();
    res.json({ id: r.lastInsertRowid });
  } catch { res.status(409).json({ erreur: "Cet identifiant existe déjà." }); }
});
app.put("/api/utilisateurs/:id", auth, admin, (req, res) => {
  const u = db.prepare("SELECT * FROM utilisateurs WHERE id=?").get(req.params.id);
  if (!u) return res.status(404).json({ erreur: "Introuvable." });
  const { nom = u.nom, poste = u.poste, actif = !!u.actif, mdp } = req.body || {};
  if (u.id === req.user.id && !actif) return res.status(400).json({ erreur: "Vous ne pouvez pas désactiver votre propre compte." });
  db.prepare("UPDATE utilisateurs SET nom=?,poste=?,actif=? WHERE id=?").run(nom, poste, actif ? 1 : 0, u.id);
  if (mdp) {
    if (mdp.length < 8) return res.status(400).json({ erreur: "Mot de passe : 8 caractères minimum." });
    db.prepare("UPDATE utilisateurs SET mdp_hash=?, doit_changer_mdp=1 WHERE id=?").run(bcrypt.hashSync(mdp, 10), u.id);
  }
  synchroniserTurso();
  res.json({ ok: true });
});
app.delete("/api/utilisateurs/:id", auth, admin, (req, res) => {
  const u = db.prepare("SELECT * FROM utilisateurs WHERE id=?").get(req.params.id);
  if (!u) return res.status(404).json({ erreur: "Introuvable." });
  if (u.id === req.user.id) return res.status(400).json({ erreur: "Vous ne pouvez pas supprimer votre propre compte administrateur." });
  db.prepare("DELETE FROM saisies WHERE user_id=?").run(u.id);
  db.prepare("DELETE FROM connexions WHERE user_id=?").run(u.id);
  db.prepare("DELETE FROM activites_pc WHERE user_id=?").run(u.id);
  db.prepare("DELETE FROM validations WHERE user_id=?").run(u.id);
  db.prepare("DELETE FROM utilisateurs WHERE id=?").run(u.id);
  synchroniserTurso();
  res.json({ ok: true });
});
app.get("/api/parametres", auth, (req, res) => res.json({ activites: JSON.parse(getParam("activites")), heure_attendue: getParam("heure_attendue") }));
app.put("/api/parametres", auth, admin, (req, res) => {
  const { activites, heure_attendue } = req.body || {};
  if (Array.isArray(activites) && activites.length) setParam("activites", JSON.stringify(activites.map(a => String(a).trim()).filter(Boolean)));
  if (/^\d{2}:\d{2}$/.test(heure_attendue || "")) setParam("heure_attendue", heure_attendue);
  res.json({ ok: true });
});
app.get("/api/tableau-de-bord", auth, admin, (req, res) => {
  const date = dateValide(req.query.date) ? req.query.date : maintenantLocal().date, mois = date.slice(0, 7), [a, b] = borne(mois);
  const att = getParam("heure_attendue"), acts = JSON.parse(getParam("activites"));
  const lignes = db.prepare("SELECT * FROM utilisateurs WHERE actif=1 AND role='employe' ORDER BY nom").all().map(u => {
    const ar = arrivee(u.id, date);
    const act = db.prepare("SELECT * FROM activites_pc WHERE user_id=? AND date=?").get(u.id, date);
    const somme = (extra, ...p) => db.prepare(`SELECT COALESCE(SUM(heures),0) h FROM saisies WHERE user_id=? ${extra}`).get(u.id, ...p).h;
    
    // Calcul inactivité en minutes
    const inactifMin = act ? Math.round((act.inactif_secondes || 0) / 60) : 0;
    
    return {
      id: u.id, nom: u.nom, arrivee: ar, statut: !ar ? "Pas encore connecté(e)" : ar > att ? "En retard" : "À l'heure",
      heures_jour: somme("AND date=?", date), heures_mois: somme("AND date BETWEEN ? AND ?", a, b),
      par_activite: Object.fromEntries(acts.map(x => [x, somme("AND date BETWEEN ? AND ? AND activite=?", a, b, x)])),
      statut_mois: statutMois(u.id, mois).statut,
      pc: act?.pc || "—",
      derniere_action: act?.derniere_action || act?.heure_dernier_signal || "—",
      inactif_minutes: inactifMin,
      heure_arret: act?.arret_detecte ? (act.heure_arret || act.heure_dernier_signal) : (act?.heure_dernier_signal ? `En service (${act.heure_dernier_signal})` : "—")
    };
  });
  res.json({ date, mois, heure_attendue: att, activites: acts, lignes });
});

// ---------- exports
const VIOLET = "6B2D8B", ORANGE = "E07B00";
app.get("/api/export/excel", auth, async (req, res) => {
  const u = cible(req); if (!u) return res.status(403).json({ erreur: "Accès refusé." });
  const annee = /^\d{4}$/.test(req.query.annee || "") ? req.query.annee : maintenantLocal().date.slice(0, 4);
  const wb = new ExcelJS.Workbook(); wb.creator = "ACH ProGestions & Outils";
  const fill = c => ({ type: "pattern", pattern: "solid", fgColor: { argb: "FF" + c } });
  MOIS.forEach((nomMois, i) => {
    const mois = `${annee}-${String(i + 1).padStart(2, "0")}`, ws = wb.addWorksheet(nomMois, { views: [{ showGridLines: false }] });
    const st = statutMois(u.id, mois), lignes = saisiesMois(u.id, mois);
    ws.columns = [{ width: 13 }, { width: 11 }, { width: 24 }, { width: 24 }, { width: 26 }, { width: 46 }, { width: 9 }];
    ws.mergeCells("A1:G1"); Object.assign(ws.getCell("A1"), { value: `ACH ProGestions & Outils — Timesheet ${nomMois} ${annee}` });
    ws.getCell("A1").font = { name: "Arial", size: 15, bold: true, color: { argb: "FFFFFFFF" } }; ws.getCell("A1").fill = fill(VIOLET); ws.getRow(1).height = 32;
    ws.getCell("A1").alignment = { vertical: "middle", indent: 1 };
    [["Nom", u.nom], ["Poste", u.poste || ""], ["Mois", `${nomMois} ${annee}`], ["Validation", st.statut + (st.statut === "Validé" && st.valide_par ? ` (${st.valide_par}, ${st.date_validation})` : "")]]
      .forEach(([k, v], r) => { ws.getCell(3 + r, 1).value = k; ws.getCell(3 + r, 1).font = { name: "Arial", bold: true, color: { argb: "FF" + VIOLET } }; ws.getCell(3 + r, 2).value = v; ws.getCell(3 + r, 2).font = { name: "Arial", bold: true, color: { argb: "FF" + ORANGE } }; });
    const h = ws.getRow(8); ["Date", "Arrivée", "Activité", "Client / Contact", "Tâche", "Description", "Heures"].forEach((t, c) => {
      const cell = h.getCell(c + 1); cell.value = t; cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } }; cell.fill = fill(VIOLET); cell.alignment = { horizontal: "center", vertical: "middle", wrap_text: true };
    }); h.height = 24;
    lignes.forEach((l, k) => {
      const [y, m, d] = l.date.split("-"), r = ws.getRow(9 + k);
      [`${d}/${m}/${y}`, l.arrivee, l.activite, l.client, l.tache, l.description, l.heures].forEach((v, c) => {
        const cell = r.getCell(c + 1); cell.value = v; cell.font = { name: "Arial", size: 10 }; cell.alignment = { vertical: "top", wrapText: c === 5 };
        cell.border = { bottom: { style: "thin", color: { argb: "FFD9C7E3" } } };
      });
    });
    const fin = 9 + lignes.length + 1;
    ws.getCell(fin, 6).value = "TOTAL HEURES"; ws.getCell(fin, 6).alignment = { horizontal: "right" };
    ws.getCell(fin, 7).value = lignes.length ? { formula: `SUM(G9:G${8 + lignes.length})` } : 0;
    [6, 7].forEach(c => { const x = ws.getCell(fin, c); x.font = { name: "Arial", bold: true, color: { argb: "FF" + ORANGE } }; x.fill = fill("FDF1E3"); });
    ws.getCell(fin + 3, 1).value = "Signature employé : ____________________"; ws.getCell(fin + 3, 5).value = "Signature direction : ____________________";
    ws.getCell(fin + 3, 1).font = ws.getCell(fin + 3, 5).font = { name: "Arial", size: 10 };
  });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="Timesheet_${u.identifiant}_${annee}.xlsx"`);
  await wb.xlsx.write(res); res.end();
});
app.get("/api/export/pdf", auth, (req, res) => {
  const u = cible(req); if (!u) return res.status(403).json({ erreur: "Accès refusé." });
  const mois = req.query.mois; if (!moisValide(mois)) return res.status(400).json({ erreur: "Mois invalide." });
  const [annee, mm] = mois.split("-"), nomMois = MOIS[Number(mm) - 1], st = statutMois(u.id, mois), lignes = saisiesMois(u.id, mois);
  const doc = new PDFDocument({ size: "A4", margin: 36 });
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="Timesheet_${u.identifiant}_${mois}.pdf"`);
  doc.pipe(res);
  const W = doc.page.width - 72;
  const entete = () => { doc.rect(36, 30, W, 40).fill("#" + VIOLET); doc.fillColor("#fff").font("Helvetica-Bold").fontSize(15).text(`ACH ProGestions & Outils — Timesheet ${nomMois} ${annee}`, 48, 44, { width: W - 24 }); doc.fillColor("#000"); };
  entete();
  doc.font("Helvetica").fontSize(10).fillColor("#" + VIOLET).text("Employé : ", 36, 84, { continued: true }).fillColor("#000").text(`${u.nom}${u.poste ? " — " + u.poste : ""}`);
  doc.fillColor("#" + VIOLET).text("Validation : ", { continued: true }).fillColor("#" + ORANGE).text(st.statut + (st.statut === "Validé" && st.valide_par ? ` (${st.valide_par}, ${st.date_validation})` : ""));
  const cols = [[54, "Date"], [38, "Arrivée"], [88, "Activité"], [84, "Client"], [84, "Tâche"], [W - 54 - 38 - 88 - 84 - 84 - 42, "Description"], [42, "Heures"]];
  const ligne = (vals, y, head) => {
    let x = 36; const h = head ? 18 : Math.max(16, ...vals.map((v, i) => doc.heightOfString(String(v), { width: cols[i][0] - 6 }) + 6));
    if (y + h > doc.page.height - 100) { doc.addPage(); entete(); y = 84; }
    if (head) doc.rect(36, y, W, h).fill("#" + VIOLET);
    vals.forEach((v, i) => { doc.fillColor(head ? "#fff" : "#000").font(head ? "Helvetica-Bold" : "Helvetica").fontSize(8.5).text(String(v), x + 3, y + 4, { width: cols[i][0] - 6, align: i === 6 ? "right" : "left" }); x += cols[i][0]; });
    if (!head) doc.moveTo(36, y + h).lineTo(36 + W, y + h).lineWidth(0.4).strokeColor("#D9C7E3").stroke();
    return y + h;
  };
  let y = ligne(cols.map(c => c[1]), 124, true), total = 0;
  lignes.forEach(l => { const [Y, M, D] = l.date.split("-"); total += l.heures; y = ligne([`${D}/${M}/${Y}`, l.arrivee, l.activite, l.client, l.tache, l.description, l.heures], y); });
  if (y > doc.page.height - 130) { doc.addPage(); entete(); y = 84; }
  doc.rect(36, y + 6, W, 20).fill("#FDF1E3"); doc.fillColor("#" + ORANGE).font("Helvetica-Bold").fontSize(10).text(`TOTAL : ${total} h`, 44, y + 12, { width: W - 16, align: "right" });
  const ys = Math.min(y + 70, doc.page.height - 90);
  doc.fillColor("#000").font("Helvetica").fontSize(9).text("Signature de l'employé :", 36, ys).text("Signature de la direction :", 320, ys);
  doc.moveTo(36, ys + 40).lineTo(230, ys + 40).strokeColor("#999").stroke(); doc.moveTo(320, ys + 40).lineTo(530, ys + 40).stroke();
  doc.end();
});

app.get("/api/sante", (req, res) => res.json({ ok: true }));
const srv = app.listen(PORT, () => console.log(`Serveur ACH démarré sur le port ${PORT}`));
srv.on("error", err => {
  if (err.code === "EADDRINUSE") {
    console.log(`Port ${PORT} déjà utilisé, serveur existant conservé.`);
  } else {
    console.error("Erreur serveur:", err);
  }
});
