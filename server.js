const express = require('express');
const Database = require('better-sqlite3');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// ── DATABASE INIT ────────────────────────────────
const dataDir = path.join(__dirname, 'data');
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir);

const db = new Database(path.join(dataDir, 'netcollect.db'));
db.pragma('journal_mode = WAL'); // Performance en lecture concurrente

db.exec(`
  CREATE TABLE IF NOT EXISTS bus (
    id          TEXT PRIMARY KEY,
    lot         TEXT    DEFAULT '',
    dor         TEXT    DEFAULT '',
    ext         TEXT    DEFAULT '',
    etat        INTEGER DEFAULT 0,
    rr          TEXT    DEFAULT '',
    av          INTEGER DEFAULT 0,
    liv         TEXT    DEFAULT '',
    mes         TEXT    DEFAULT '',
    risque      TEXT    DEFAULT 'Faible',
    comment     TEXT    DEFAULT '',
    created_at  TEXT    DEFAULT (datetime('now')),
    updated_at  TEXT    DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS delivery (
    topo        TEXT PRIMARY KEY,
    lot         TEXT DEFAULT '',
    hRcv        TEXT DEFAULT 'NON',
    hF          TEXT DEFAULT 'NON',
    hRnv        TEXT DEFAULT 'NON',
    cRcv        TEXT DEFAULT 'NON',
    cF          TEXT DEFAULT 'NON',
    cRnv        TEXT DEFAULT 'NON',
    vO          TEXT DEFAULT 'NON',
    cf          TEXT DEFAULT 'NON',
    updated_at  TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS mesures (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    topo        TEXT NOT NULL,
    rr          TEXT DEFAULT '',
    nbT         INTEGER DEFAULT 0,
    moe         TEXT DEFAULT 'Nok',
    nbM         INTEGER DEFAULT 0,
    ok          INTEGER DEFAULT 0,
    nok         INTEGER DEFAULT 0,
    updated_at  TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS cdd (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    lot         TEXT DEFAULT '',
    topo        TEXT DEFAULT '',
    dept        TEXT DEFAULT '',
    rr          TEXT DEFAULT '',
    sc          TEXT DEFAULT '',
    ns          TEXT DEFAULT '',
    mad         TEXT DEFAULT '',
    cmd         TEXT DEFAULT '',
    pfto        TEXT DEFAULT 'NON',
    sin3        TEXT DEFAULT 'NON',
    comment     TEXT DEFAULT '',
    updated_at  TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS emplacements (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    topo        TEXT DEFAULT '',
    lot         TEXT DEFAULT '',
    site        TEXT DEFAULT '',
    dept        TEXT DEFAULT '',
    rr          TEXT DEFAULT '',
    ep          TEXT DEFAULT 'NON',
    rdv         TEXT DEFAULT 'NON',
    aB          TEXT DEFAULT 'NON',
    inst        TEXT DEFAULT 'NON',
    cons        TEXT DEFAULT 'NON',
    aC          TEXT DEFAULT 'NON',
    updated_at  TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS history (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    entity      TEXT NOT NULL,
    entity_id   TEXT NOT NULL,
    field       TEXT NOT NULL,
    old_val     TEXT,
    new_val     TEXT,
    user_name   TEXT DEFAULT 'système',
    created_at  TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS comments (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    topo        TEXT NOT NULL,
    author      TEXT DEFAULT 'Équipe',
    body        TEXT NOT NULL,
    created_at  TEXT DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_history_entity ON history(entity_id);
  CREATE INDEX IF NOT EXISTS idx_comments_topo ON comments(topo);
`);

// ── MIDDLEWARE ───────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ── HELPERS ──────────────────────────────────────
function logHistory(entity, entityId, field, oldVal, newVal, user = 'système') {
  if (String(oldVal ?? '') !== String(newVal ?? '')) {
    db.prepare(
      'INSERT INTO history (entity, entity_id, field, old_val, new_val, user_name) VALUES (?,?,?,?,?,?)'
    ).run(entity, entityId, field, String(oldVal ?? ''), String(newVal ?? ''), user);
  }
}

// ── BUS ROUTES ───────────────────────────────────
app.get('/api/bus', (req, res) => {
  const rows = db.prepare('SELECT * FROM bus ORDER BY lot, id').all();
  res.json(rows);
});

app.get('/api/bus/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM bus WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'BUS introuvable' });
  res.json(row);
});

app.post('/api/bus', (req, res) => {
  const b = req.body;
  try {
    db.prepare(`
      INSERT INTO bus (id, lot, dor, ext, etat, rr, av, liv, mes, risque, comment)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(b.id, b.lot||'', b.dor||'', b.ext||'', b.etat??0, b.rr||'', b.av??0, b.liv||'', b.mes||'', b.risque||'Faible', b.comment||'');
    res.status(201).json(db.prepare('SELECT * FROM bus WHERE id = ?').get(b.id));
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.put('/api/bus/:id', (req, res) => {
  const { id } = req.params;
  const fields = req.body;
  const current = db.prepare('SELECT * FROM bus WHERE id = ?').get(id);
  if (!current) return res.status(404).json({ error: 'BUS introuvable' });

  const allowed = ['lot','dor','ext','etat','rr','av','liv','mes','risque','comment'];
  const keys = Object.keys(fields).filter(k => allowed.includes(k));
  if (!keys.length) return res.status(400).json({ error: 'Aucun champ valide' });

  // Historique
  const user = fields.user || 'système';
  keys.forEach(k => logHistory('bus', id, k, current[k], fields[k], user));

  const setClause = keys.map(k => `${k} = ?`).join(', ');
  db.prepare(`UPDATE bus SET ${setClause}, updated_at = datetime('now') WHERE id = ?`)
    .run(...keys.map(k => fields[k]), id);

  res.json(db.prepare('SELECT * FROM bus WHERE id = ?').get(id));
});

// ── DELIVERY ROUTES ──────────────────────────────
app.get('/api/delivery', (req, res) => {
  res.json(db.prepare('SELECT * FROM delivery ORDER BY lot, topo').all());
});

app.put('/api/delivery/:topo', (req, res) => {
  const { topo } = req.params;
  const f = req.body;
  const allowed = ['lot','hRcv','hF','hRnv','cRcv','cF','cRnv','vO','cf'];
  const current = db.prepare('SELECT * FROM delivery WHERE topo = ?').get(topo);

  db.prepare(`
    INSERT INTO delivery (topo, lot, hRcv, hF, hRnv, cRcv, cF, cRnv, vO, cf)
    VALUES (?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(topo) DO UPDATE SET
      lot=excluded.lot, hRcv=excluded.hRcv, hF=excluded.hF, hRnv=excluded.hRnv,
      cRcv=excluded.cRcv, cF=excluded.cF, cRnv=excluded.cRnv, vO=excluded.vO,
      cf=excluded.cf, updated_at=datetime('now')
  `).run(topo, f.lot||'', f.hRcv||'NON', f.hF||'NON', f.hRnv||'NON',
         f.cRcv||'NON', f.cF||'NON', f.cRnv||'NON', f.vO||'NON', f.cf||'NON');

  if (current) {
    allowed.forEach(k => { if (f[k] !== undefined) logHistory('delivery', topo, k, current[k], f[k]); });
  }
  res.json({ ok: true });
});

// ── MESURES ROUTES ───────────────────────────────
app.get('/api/mesures', (req, res) => {
  res.json(db.prepare('SELECT * FROM mesures ORDER BY topo').all());
});

app.put('/api/mesures/:id', (req, res) => {
  const f = req.body;
  db.prepare(`
    UPDATE mesures SET rr=?, nbT=?, moe=?, nbM=?, ok=?, nok=?, updated_at=datetime('now')
    WHERE id=?
  `).run(f.rr||'', f.nbT||0, f.moe||'Nok', f.nbM||0, f.ok||0, f.nok||0, req.params.id);
  res.json({ ok: true });
});

// ── CDD ROUTES ───────────────────────────────────
app.get('/api/cdd', (req, res) => {
  res.json(db.prepare('SELECT * FROM cdd ORDER BY lot, topo').all());
});

app.put('/api/cdd/:id', (req, res) => {
  const f = req.body;
  const current = db.prepare('SELECT * FROM cdd WHERE id = ?').get(req.params.id);
  db.prepare(`
    UPDATE cdd SET pfto=?, sin3=?, comment=?, updated_at=datetime('now') WHERE id=?
  `).run(f.pfto||'NON', f.sin3||'NON', f.comment||'', req.params.id);
  if (current) {
    ['pfto','sin3','comment'].forEach(k => { if(f[k]!==undefined) logHistory('cdd', String(req.params.id), k, current[k], f[k]); });
  }
  res.json({ ok: true });
});

// ── EMPLACEMENTS ROUTES ──────────────────────────
app.get('/api/emplacements', (req, res) => {
  res.json(db.prepare('SELECT * FROM emplacements ORDER BY lot, topo').all());
});

app.put('/api/emplacements/:id', (req, res) => {
  const f = req.body;
  db.prepare(`
    UPDATE emplacements SET ep=?, rdv=?, aB=?, inst=?, cons=?, aC=?, updated_at=datetime('now') WHERE id=?
  `).run(f.ep||'NON', f.rdv||'NON', f.aB||'NON', f.inst||'NON', f.cons||'NON', f.aC||'NON', req.params.id);
  res.json({ ok: true });
});

// ── HISTORY & COMMENTS ───────────────────────────
app.get('/api/history/:id', (req, res) => {
  const rows = db.prepare(
    'SELECT * FROM history WHERE entity_id = ? ORDER BY created_at DESC LIMIT 100'
  ).all(req.params.id);
  res.json(rows);
});

app.get('/api/comments/:topo', (req, res) => {
  res.json(db.prepare('SELECT * FROM comments WHERE topo = ? ORDER BY created_at DESC').all(req.params.topo));
});

app.post('/api/comments/:topo', (req, res) => {
  const { body: text, author } = req.body;
  if (!text?.trim()) return res.status(400).json({ error: 'Commentaire vide' });
  const r = db.prepare('INSERT INTO comments (topo, author, body) VALUES (?,?,?)').run(
    req.params.topo, author || 'Équipe', text.trim()
  );
  res.status(201).json({ id: r.lastInsertRowid, topo: req.params.topo, body: text, author: author||'Équipe' });
});

// ── IMPORT EXCEL ─────────────────────────────────
app.post('/api/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Aucun fichier reçu' });

  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const result = { imported: 0, updated: 0, sheets: {}, warnings: [], errors: [] };

    // ── Onglet Etat BUS ──
    const bsSheet = wb.Sheets['Etat BUS'];
    if (bsSheet) {
      const rows = XLSX.utils.sheet_to_json(bsSheet, { defval: '' });
      result.sheets['Etat BUS'] = rows.length;

      const upsert = db.prepare(`
        INSERT INTO bus (id, lot, dor, ext, etat)
        VALUES (?,?,?,?,?)
        ON CONFLICT(id) DO UPDATE SET
          lot=excluded.lot, dor=excluded.dor, ext=excluded.ext,
          etat=excluded.etat, updated_at=datetime('now')
      `);
      const tx = db.transaction((rows) => {
        rows.forEach((row, idx) => {
          const topo = String(row['Nom topologie'] || '').trim();
          if (!topo) { result.warnings.push(`Ligne ${idx + 2}: topologie manquante`); return; }
          const es = String(row['Etat BUS'] || '').trim();
          const en = parseInt(es.charAt(0));
          const existing = db.prepare('SELECT id FROM bus WHERE id = ?').get(topo);
          upsert.run(topo, String(row['Lot']||''), String(row['DOR']||''), String(row['Extrémité']||row['Extrémités']||''), isNaN(en) ? 0 : en);
          if (existing) result.updated++; else result.imported++;
        });
      });
      tx(rows);
    } else {
      result.errors.push("Onglet 'Etat BUS' introuvable");
    }

    // ── Onglet HME (dates) ──
    const hmeSheet = wb.Sheets['HME'];
    if (hmeSheet) {
      const rows = XLSX.utils.sheet_to_json(hmeSheet, { defval: '' });
      result.sheets['HME'] = rows.length;
      rows.forEach(r => {
        const topo = String(r['ReferenceRegroupement'] || '').trim();
        if (!topo) return;
        if (r['Semaine Prév. Livraison']) db.prepare("UPDATE bus SET liv=? WHERE id=?").run(String(r['Semaine Prév. Livraison']), topo);
        if (r['Semaine Prév. MES']) db.prepare("UPDATE bus SET mes=? WHERE id=?").run(String(r['Semaine Prév. MES']), topo);
      });
    }

    // ── Onglet Suivi TRAVAUX (RR + avancement) ──
    const trxSheet = wb.Sheets['Suivi TRAVAUX'];
    if (trxSheet) {
      const rows = XLSX.utils.sheet_to_json(trxSheet, { defval: '' });
      result.sheets['Suivi TRAVAUX'] = rows.length;
      rows.forEach(r => {
        const topo = String(r['Typologie'] || r['Topologie'] || '').trim();
        if (!topo) return;
        if (r['RR']) db.prepare("UPDATE bus SET rr=? WHERE id=?").run(String(r['RR']).trim(), topo);
        if (r['Avancement BUS'] !== undefined && r['Avancement BUS'] !== '') {
          const av = Math.round(parseFloat(r['Avancement BUS']) * 100);
          if (!isNaN(av)) db.prepare("UPDATE bus SET av=? WHERE id=?").run(av, topo);
        }
      });
    }

    // ── Onglet SUIVI DELIVERY ──
    const dlvSheet = wb.Sheets['SUIVI DELIVERY'];
    if (dlvSheet) {
      const rows = XLSX.utils.sheet_to_json(dlvSheet, { defval: '' });
      result.sheets['SUIVI DELIVERY'] = rows.length;
      const upsert = db.prepare(`
        INSERT INTO delivery (topo, lot, hRcv, hF, hRnv, cRcv, cF, cRnv, vO, cf)
        VALUES (?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(topo) DO UPDATE SET
          hRcv=excluded.hRcv, hF=excluded.hF, hRnv=excluded.hRnv,
          cRcv=excluded.cRcv, cF=excluded.cF, cRnv=excluded.cRnv,
          vO=excluded.vO, cf=excluded.cf, updated_at=datetime('now')
      `);
      rows.forEach(r => {
        const topo = String(r['Nom topologie'] || r['Topologie'] || '').trim();
        if (!topo) return;
        const yn = v => (v && String(v).toUpperCase().includes('OUI')) ? 'OUI' : 'NON';
        upsert.run(topo, String(r['Lot']||''), yn(r['HEPOC transmis Orange']), yn(r['HEPOC complété Free']),
          yn(r['HEPOC renvoyé Orange']), yn(r['CPM transmis Orange']), yn(r['CPM complété Free']),
          yn(r['CPM renvoyé Orange']), yn(r['Validation Orange']), yn(r['Commande ferme']));
      });
    }

    // ── Onglet Mesures ──
    const mesSheet = wb.Sheets['Mesures'];
    if (mesSheet) {
      const rows = XLSX.utils.sheet_to_json(mesSheet, { defval: '' });
      result.sheets['Mesures'] = rows.length;
    }

    // ── Onglet CDD ──
    const cddSheet = wb.Sheets['CDD'];
    if (cddSheet) {
      const rows = XLSX.utils.sheet_to_json(cddSheet, { defval: '' });
      result.sheets['CDD'] = rows.length;
    }

    result.warnings = result.warnings.slice(0, 10);
    res.json(result);

  } catch (e) {
    console.error('Import error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── STATS (dashboard) ────────────────────────────
app.get('/api/stats', (req, res) => {
  const byEtat = db.prepare(`
    SELECT etat, COUNT(*) as count FROM bus GROUP BY etat
  `).all();
  const byLot = db.prepare(`
    SELECT lot, COUNT(*) as total,
      SUM(CASE WHEN etat=4 THEN 1 ELSE 0 END) as trx,
      SUM(CASE WHEN etat>=5 THEN 1 ELSE 0 END) as livres,
      ROUND(AVG(av)) as avMoy
    FROM bus GROUP BY lot
  `).all();
  const avGlobal = db.prepare('SELECT ROUND(AVG(av)) as avg FROM bus').get();
  const cfCount = db.prepare("SELECT COUNT(*) as n FROM delivery WHERE cf='OUI'").get();
  const nokCount = db.prepare('SELECT SUM(nok) as n FROM mesures').get();
  const sin3Nok = db.prepare("SELECT COUNT(*) as n FROM cdd WHERE sin3='NON'").get();
  const riskHigh = db.prepare("SELECT COUNT(*) as n FROM bus WHERE risque='Élevé'").get();

  res.json({ byEtat, byLot, avGlobal: avGlobal.avg || 0, cfCount: cfCount.n, nokCount: nokCount.n || 0, sin3Nok: sin3Nok.n, riskHigh: riskHigh.n });
});

// ── CATCH ALL → SPA ──────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── START ─────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅ NetCollect démarré sur http://localhost:${PORT}`);
  console.log(`📁 Base de données : ${path.join(dataDir, 'netcollect.db')}\n`);
});
