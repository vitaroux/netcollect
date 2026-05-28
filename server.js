const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── JSON FILE STORAGE ─────────────────────────────
const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const TABLES = ['bus','delivery','mesures','cdd','emplacements','history','comments'];

function loadTable(name) {
  const file = path.join(DATA_DIR, name + '.json');
  if (!fs.existsSync(file)) return [];
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); }
  catch(e) { return []; }
}

function saveTable(name, data) {
  const file = path.join(DATA_DIR, name + '.json');
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

function getDB() {
  const db = {};
  TABLES.forEach(t => db[t] = loadTable(t));
  return db;
}

function now() { return new Date().toISOString().replace('T',' ').slice(0,19); }

function logHistory(db, entity, entityId, field, oldVal, newVal, user='système') {
  if (String(oldVal ?? '') !== String(newVal ?? '')) {
    db.history.push({ id: Date.now(), entity, entity_id: entityId, field,
      old_val: String(oldVal ?? ''), new_val: String(newVal ?? ''),
      user_name: user, created_at: now() });
    saveTable('history', db.history);
  }
}

// ── MIDDLEWARE ────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50*1024*1024 } });

// ── BUS ───────────────────────────────────────────
app.get('/api/bus', (req, res) => {
  const db = getDB();
  res.json(db.bus.sort((a,b) => (a.lot+a.id).localeCompare(b.lot+b.id)));
});

app.get('/api/bus/:id', (req, res) => {
  const db = getDB();
  const row = db.bus.find(b => b.id === req.params.id);
  if (!row) return res.status(404).json({ error: 'BUS introuvable' });
  res.json(row);
});

app.post('/api/bus', (req, res) => {
  const db = getDB();
  const b  = req.body;
  if (db.bus.find(x => x.id === b.id))
    return res.status(400).json({ error: 'BUS déjà existant' });
  const row = { id:b.id, lot:b.lot||'', dor:b.dor||'', ext:b.ext||'',
    etat:b.etat??0, rr:b.rr||'', av:b.av??0, liv:b.liv||'', mes:b.mes||'',
    risque:b.risque||'Faible', comment:b.comment||'',
    created_at:now(), updated_at:now() };
  db.bus.push(row);
  saveTable('bus', db.bus);
  res.status(201).json(row);
});

app.put('/api/bus/:id', (req, res) => {
  const db = getDB();
  const idx = db.bus.findIndex(b => b.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'BUS introuvable' });
  const allowed = ['lot','dor','ext','etat','rr','av','liv','mes','risque','comment'];
  const user = req.body.user || 'système';
  allowed.forEach(k => {
    if (req.body[k] !== undefined) {
      logHistory(db, 'bus', req.params.id, k, db.bus[idx][k], req.body[k], user);
      db.bus[idx][k] = req.body[k];
    }
  });
  db.bus[idx].updated_at = now();
  saveTable('bus', db.bus);
  res.json(db.bus[idx]);
});

// ── DELIVERY ──────────────────────────────────────
app.get('/api/delivery', (req, res) => {
  res.json(loadTable('delivery'));
});

app.put('/api/delivery/:topo', (req, res) => {
  const db = getDB();
  const f  = req.body;
  const idx = db.delivery.findIndex(d => d.topo === req.params.topo);
  const row = { topo:req.params.topo, lot:f.lot||'',
    hRcv:f.hRcv||'NON', hF:f.hF||'NON', hRnv:f.hRnv||'NON',
    cRcv:f.cRcv||'NON', cF:f.cF||'NON', cRnv:f.cRnv||'NON',
    vO:f.vO||'NON', cf:f.cf||'NON', updated_at:now() };
  if (idx >= 0) db.delivery[idx] = row;
  else db.delivery.push(row);
  saveTable('delivery', db.delivery);
  res.json({ ok: true });
});

// ── MESURES ───────────────────────────────────────
app.get('/api/mesures', (req, res) => res.json(loadTable('mesures')));

app.put('/api/mesures/:id', (req, res) => {
  const db = getDB();
  const f  = req.body;
  const idx = db.mesures.findIndex(m => m.id === +req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Mesure introuvable' });
  Object.assign(db.mesures[idx], { rr:f.rr||'', nbT:f.nbT||0, moe:f.moe||'Nok',
    nbM:f.nbM||0, ok:f.ok||0, nok:f.nok||0, updated_at:now() });
  saveTable('mesures', db.mesures);
  res.json({ ok: true });
});

// ── CDD ───────────────────────────────────────────
app.get('/api/cdd', (req, res) => res.json(loadTable('cdd')));

app.put('/api/cdd/:id', (req, res) => {
  const db = getDB();
  const f  = req.body;
  const idx = db.cdd.findIndex(c => c.id === +req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'CDD introuvable' });
  const user = f.user || 'système';
  ['pfto','sin3','comment'].forEach(k => {
    if (f[k] !== undefined) {
      logHistory(db, 'cdd', String(req.params.id), k, db.cdd[idx][k], f[k], user);
      db.cdd[idx][k] = f[k];
    }
  });
  db.cdd[idx].updated_at = now();
  saveTable('cdd', db.cdd);
  res.json({ ok: true });
});

// ── EMPLACEMENTS ──────────────────────────────────
app.get('/api/emplacements', (req, res) => res.json(loadTable('emplacements')));

app.put('/api/emplacements/:id', (req, res) => {
  const db = getDB();
  const f  = req.body;
  const idx = db.emplacements.findIndex(e => e.id === +req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Emplacement introuvable' });
  Object.assign(db.emplacements[idx], { ep:f.ep||'NON', rdv:f.rdv||'NON',
    aB:f.aB||'NON', inst:f.inst||'NON', cons:f.cons||'NON', aC:f.aC||'NON',
    updated_at:now() });
  saveTable('emplacements', db.emplacements);
  res.json({ ok: true });
});

// ── HISTORY & COMMENTS ────────────────────────────
app.get('/api/history/:id', (req, res) => {
  const h = loadTable('history')
    .filter(r => r.entity_id === req.params.id)
    .sort((a,b) => b.created_at.localeCompare(a.created_at))
    .slice(0, 100);
  res.json(h);
});

app.get('/api/comments/:topo', (req, res) => {
  const c = loadTable('comments')
    .filter(r => r.topo === req.params.topo)
    .sort((a,b) => b.created_at.localeCompare(a.created_at));
  res.json(c);
});

app.post('/api/comments/:topo', (req, res) => {
  const { body: text, author } = req.body;
  if (!text?.trim()) return res.status(400).json({ error: 'Commentaire vide' });
  const db = getDB();
  const row = { id: Date.now(), topo: req.params.topo,
    author: author||'Équipe', body: text.trim(), created_at: now() };
  db.comments.push(row);
  saveTable('comments', db.comments);
  res.status(201).json(row);
});

// ── STATS ─────────────────────────────────────────
app.get('/api/stats', (req, res) => {
  const bus = loadTable('bus');
  const mes = loadTable('mesures');
  const cdd = loadTable('cdd');
  const dlv = loadTable('delivery');

  const byEtat = [0,1,2,3,4,5,6].map(e => ({
    etat: e, count: bus.filter(b => b.etat === e).length
  })).filter(r => r.count > 0);

  const lots = [...new Set(bus.map(b => b.lot).filter(Boolean))];
  const byLot = lots.map(lot => {
    const lb = bus.filter(b => b.lot === lot);
    return { lot, total: lb.length,
      trx: lb.filter(b => b.etat === 4).length,
      livres: lb.filter(b => b.etat >= 5).length,
      avMoy: lb.length ? Math.round(lb.reduce((a,b)=>a+b.av,0)/lb.length) : 0 };
  });

  const avGlobal  = bus.length ? Math.round(bus.reduce((a,b)=>a+b.av,0)/bus.length) : 0;
  const cfCount   = dlv.filter(d => d.cf === 'OUI').length;
  const nokCount  = mes.reduce((a,m) => a + (m.nok||0), 0);
  const sin3Nok   = cdd.filter(c => c.sin3 === 'NON').length;
  const riskHigh  = bus.filter(b => b.risque === 'Élevé').length;

  res.json({ byEtat, byLot, avGlobal, cfCount, nokCount, sin3Nok, riskHigh });
});

// ── IMPORT EXCEL ──────────────────────────────────
app.post('/api/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Aucun fichier reçu' });
  try {
    const wb     = XLSX.read(req.file.buffer, { type: 'buffer' });
    const db     = getDB();
    const result = { imported:0, updated:0, sheets:{}, warnings:[], errors:[] };

    // Etat BUS
    const bsSheet = wb.Sheets['Etat BUS'];
    if (bsSheet) {
      const rows = XLSX.utils.sheet_to_json(bsSheet, { defval: '' });
      result.sheets['Etat BUS'] = rows.length;
      rows.forEach((row, idx) => {
        const topo = String(row['Nom topologie'] || '').trim();
        if (!topo) { result.warnings.push(`Ligne ${idx+2}: topologie manquante`); return; }
        const es = String(row['Etat BUS'] || '').trim();
        const en = parseInt(es.charAt(0));
        const existing = db.bus.findIndex(b => b.id === topo);
        const data = { id:topo, lot:String(row['Lot']||''), dor:String(row['DOR']||''),
          ext:String(row['Extrémité']||row['Extrémités']||''),
          etat:isNaN(en)?0:en, updated_at:now() };
        if (existing >= 0) {
          Object.assign(db.bus[existing], data);
          result.updated++;
        } else {
          db.bus.push({ ...data, rr:'', av:0, liv:'', mes:'', risque:'Faible',
            comment:'', created_at:now() });
          result.imported++;
        }
      });
      saveTable('bus', db.bus);
    } else {
      result.errors.push("Onglet 'Etat BUS' introuvable");
    }

    // HME (dates prévisionnelles)
    const hme = wb.Sheets['HME'];
    if (hme) {
      const rows = XLSX.utils.sheet_to_json(hme, { defval: '' });
      result.sheets['HME'] = rows.length;
      rows.forEach(r => {
        const topo = String(r['ReferenceRegroupement'] || '').trim();
        const idx  = db.bus.findIndex(b => b.id === topo);
        if (idx < 0) return;
        if (r['Semaine Prév. Livraison']) db.bus[idx].liv = String(r['Semaine Prév. Livraison']);
        if (r['Semaine Prév. MES'])       db.bus[idx].mes = String(r['Semaine Prév. MES']);
      });
      saveTable('bus', db.bus);
    }

    // Suivi TRAVAUX (RR + avancement)
    const trx = wb.Sheets['Suivi TRAVAUX'];
    if (trx) {
      const rows = XLSX.utils.sheet_to_json(trx, { defval: '' });
      result.sheets['Suivi TRAVAUX'] = rows.length;
      rows.forEach(r => {
        const topo = String(r['Typologie'] || r['Topologie'] || '').trim();
        const idx  = db.bus.findIndex(b => b.id === topo);
        if (idx < 0) return;
        if (r['RR']) db.bus[idx].rr = String(r['RR']).trim();
        if (r['Avancement BUS'] !== undefined && r['Avancement BUS'] !== '') {
          const av = Math.round(parseFloat(r['Avancement BUS']) * 100);
          if (!isNaN(av)) db.bus[idx].av = av;
        }
      });
      saveTable('bus', db.bus);
    }

    // SUIVI DELIVERY
    const dlv = wb.Sheets['SUIVI DELIVERY'];
    if (dlv) {
      const rows = XLSX.utils.sheet_to_json(dlv, { defval: '' });
      result.sheets['SUIVI DELIVERY'] = rows.length;
      const yn = v => (v && String(v).toUpperCase().includes('OUI')) ? 'OUI' : 'NON';
      rows.forEach(r => {
        const topo = String(r['Nom topologie'] || r['Topologie'] || '').trim();
        if (!topo) return;
        const row  = { topo, lot:String(r['Lot']||''),
          hRcv:yn(r['HEPOC transmis Orange']),   hF:yn(r['HEPOC complété Free']),
          hRnv:yn(r['HEPOC renvoyé Orange']),     cRcv:yn(r['CPM transmis Orange']),
          cF:yn(r['CPM complété Free']),           cRnv:yn(r['CPM renvoyé Orange']),
          vO:yn(r['Validation Orange']),           cf:yn(r['Commande ferme']),
          updated_at:now() };
        const idx = db.delivery.findIndex(d => d.topo === topo);
        if (idx >= 0) db.delivery[idx] = row; else db.delivery.push(row);
      });
      saveTable('delivery', db.delivery);
    }

    // Mesures
    const mes = wb.Sheets['Mesures'];
    if (mes) result.sheets['Mesures'] = XLSX.utils.sheet_to_json(mes,{defval:''}).length;

    // CDD
    const cdd = wb.Sheets['CDD'];
    if (cdd) result.sheets['CDD'] = XLSX.utils.sheet_to_json(cdd,{defval:''}).length;

    result.warnings = result.warnings.slice(0, 10);
    res.json(result);

  } catch(e) {
    console.error('Import error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── HEALTH CHECK ──────────────────────────────────
app.get('/health', (req, res) => res.json({ status:'ok', time:now() }));

// ── SPA FALLBACK ──────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── START ─────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅  NetCollect démarré → http://localhost:${PORT}`);
  console.log(`📁  Données            → ${DATA_DIR}\n`);
});
