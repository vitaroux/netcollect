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
  fs.writeFileSync(path.join(DATA_DIR, name + '.json'), JSON.stringify(data, null, 2));
}
function getDB() {
  const db = {};
  TABLES.forEach(t => db[t] = loadTable(t));
  return db;
}
function now() { return new Date().toISOString().replace('T',' ').slice(0,19); }
function logHistory(db, entity, entityId, field, oldVal, newVal, user='système') {
  if (String(oldVal ?? '') !== String(newVal ?? '')) {
    db.history.push({ id: Date.now() + Math.random(), entity, entity_id: entityId,
      field, old_val: String(oldVal ?? ''), new_val: String(newVal ?? ''),
      user_name: user, created_at: now() });
    saveTable('history', db.history);
  }
}

// ── HELPERS IMPORT ────────────────────────────────
// Convertit "4 - Travaux" / "En Travaux" / "Travaux" → chiffre 0-6
function parseEtat(val) {
  if (val === null || val === undefined) return 0;
  const s = String(val).trim().toLowerCase();
  if (s.startsWith('0') || s.includes('non dém') || s.includes('non dem')) return 0;
  if (s.startsWith('1') || s === 'design' || s.includes('design')) return 1;
  if (s.startsWith('2') || s === 'delivery' || s.includes('delivery')) return 2;
  if (s.startsWith('3') || s.includes('plan') || s.includes('étude') || s.includes('etude') || s.includes('analyse')) return 3;
  if (s.startsWith('4') || s.includes('travaux') || s.includes('en trav')) return 4;
  if (s.startsWith('5') || s === 'livré' || s === 'livre' || s.includes('livr')) return 5;
  if (s.startsWith('6') || s.includes('mes') || s.includes('exploit')) return 6;
  return 0;
}
// Normalise lot : "Lot1" → "LOT1", "LOT 2" → "LOT2"
function parseLot(val) {
  if (!val) return '';
  return String(val).trim().toUpperCase().replace(/\s+/g,'').replace('LOT','LOT');
}
// Nettoie les cellules avec \xa0 ou espaces
function clean(val) {
  if (val === null || val === undefined) return '';
  return String(val).replace(/\xa0/g,' ').trim();
}
// Détecte OUI / NON dans une cellule
function parseOui(val) {
  if (!val) return 'NON';
  const s = String(val).trim().toUpperCase();
  if (s === 'OUI' || s === 'O' || s === 'YES' || s === 'OK' || s.startsWith('OUI')) return 'OUI';
  return 'NON';
}

// ── MIDDLEWARE ────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50*1024*1024 } });

// ── BUS ───────────────────────────────────────────
app.get('/api/bus', (req, res) => {
  res.json(loadTable('bus').sort((a,b) => (a.lot+a.id).localeCompare(b.lot+b.id)));
});
app.get('/api/bus/:id', (req, res) => {
  const row = loadTable('bus').find(b => b.id === req.params.id);
  if (!row) return res.status(404).json({ error: 'BUS introuvable' });
  res.json(row);
});
app.post('/api/bus', (req, res) => {
  const db = getDB();
  if (db.bus.find(x => x.id === req.body.id))
    return res.status(400).json({ error: 'BUS déjà existant' });
  const b = req.body;
  const row = { id:b.id, lot:b.lot||'', dor:b.dor||'', ext:b.ext||'',
    etat:b.etat??0, rr:b.rr||'', av:b.av??0, liv:b.liv||'', mes:b.mes||'',
    risque:b.risque||'Faible', comment:b.comment||'', created_at:now(), updated_at:now() };
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
app.get('/api/delivery', (req, res) => res.json(loadTable('delivery')));
app.put('/api/delivery/:topo', (req, res) => {
  const db = getDB();
  const f  = req.body;
  const row = { topo:req.params.topo, lot:f.lot||'',
    hRcv:f.hRcv||'NON', hF:f.hF||'NON', hRnv:f.hRnv||'NON',
    cRcv:f.cRcv||'NON', cF:f.cF||'NON', cRnv:f.cRnv||'NON',
    vO:f.vO||'NON', cf:f.cf||'NON', etat:f.etat||'', updated_at:now() };
  const idx = db.delivery.findIndex(d => d.topo === req.params.topo);
  if (idx >= 0) db.delivery[idx] = row; else db.delivery.push(row);
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
      logHistory(db,'cdd',String(req.params.id),k,db.cdd[idx][k],f[k],user);
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
  res.json(loadTable('history')
    .filter(r => r.entity_id === req.params.id)
    .sort((a,b) => b.created_at.localeCompare(a.created_at)).slice(0,100));
});
app.get('/api/comments/:topo', (req, res) => {
  res.json(loadTable('comments')
    .filter(r => r.topo === req.params.topo)
    .sort((a,b) => b.created_at.localeCompare(a.created_at)));
});
app.post('/api/comments/:topo', (req, res) => {
  const { body: text, author } = req.body;
  if (!text?.trim()) return res.status(400).json({ error: 'Commentaire vide' });
  const db  = getDB();
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
    etat:e, count:bus.filter(b=>b.etat===e).length })).filter(r=>r.count>0);
  const lots = [...new Set(bus.map(b=>b.lot).filter(Boolean))].sort();
  const byLot = lots.map(lot => {
    const lb = bus.filter(b=>b.lot===lot);
    return { lot, total:lb.length, trx:lb.filter(b=>b.etat===4).length,
      livres:lb.filter(b=>b.etat>=5).length,
      avMoy:lb.length?Math.round(lb.reduce((a,b)=>a+b.av,0)/lb.length):0 };
  });
  res.json({
    byEtat, byLot,
    avGlobal: bus.length ? Math.round(bus.reduce((a,b)=>a+b.av,0)/bus.length) : 0,
    cfCount:  dlv.filter(d=>d.cf==='OUI').length,
    nokCount: mes.reduce((a,m)=>a+(m.nok||0),0),
    sin3Nok:  cdd.filter(c=>c.sin3==='NON').length,
    riskHigh: bus.filter(b=>b.risque==='Élevé').length,
  });
});

// ── IMPORT EXCEL ──────────────────────────────────
// Basé sur la VRAIE structure de SUIVI_PROJET_Collecte.xlsx
app.post('/api/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Aucun fichier reçu' });
  try {
    const wb     = XLSX.read(req.file.buffer, { type:'buffer' });
    const db     = getDB();
    const result = { imported:0, updated:0, sheets:{}, warnings:[], errors:[] };

    // ── 1. Etat BUS ──────────────────────────────
    // Colonnes: Lot | DOR | Nom topologie | Extrémité | Etat BUS
    // Ligne 1 = headers, données à partir de ligne 2
    // Lots : "Lot1", "Lot2", "Lot3"
    const bsSheet = wb.Sheets['Etat BUS'];
    if (bsSheet) {
      const rows = XLSX.utils.sheet_to_json(bsSheet, { defval: null });
      result.sheets['Etat BUS'] = rows.length;
      rows.forEach((row, idx) => {
        const topo = clean(row['Nom topologie']);
        if (!topo) { result.warnings.push(`Etat BUS ligne ${idx+2}: topologie vide`); return; }
        const etat = parseEtat(row['Etat BUS']);
        const lot  = parseLot(row['Lot']);
        const dor  = clean(row['DOR']);
        const ext  = clean(row['Extrémité']);
        const existing = db.bus.findIndex(b => b.id === topo);
        if (existing >= 0) {
          // Mise à jour — on ne touche pas aux champs enrichis manuellement
          db.bus[existing].lot  = lot  || db.bus[existing].lot;
          db.bus[existing].dor  = dor  || db.bus[existing].dor;
          db.bus[existing].ext  = ext  || db.bus[existing].ext;
          db.bus[existing].etat = etat;
          db.bus[existing].updated_at = now();
          result.updated++;
        } else {
          db.bus.push({ id:topo, lot, dor, ext, etat, rr:'', av:0,
            liv:'', mes:'', risque:'Faible', comment:'',
            created_at:now(), updated_at:now() });
          result.imported++;
        }
      });
      saveTable('bus', db.bus);
    } else {
      result.errors.push("Onglet 'Etat BUS' introuvable");
    }

    // ── 2. HME ───────────────────────────────────
    // Colonnes: ReferenceRegroupement | LOT | DROR | Extrrémités | Etat | Semaine Prév. Livraison | Semaine Prév. MES
    // Contient aussi l'état (plus à jour que Etat BUS parfois)
    const hmeSheet = wb.Sheets['HME'];
    if (hmeSheet) {
      const rows = XLSX.utils.sheet_to_json(hmeSheet, { defval: null });
      result.sheets['HME'] = rows.length;
      rows.forEach(r => {
        const topo = clean(r['ReferenceRegroupement']);
        if (!topo) return;
        const idx = db.bus.findIndex(b => b.id === topo);
        if (idx < 0) return;
        // Dates prévisionnelles
        if (r['Semaine Prév. Livraison']) db.bus[idx].liv = clean(r['Semaine Prév. Livraison']);
        if (r['Semaine Prév. MES'])       db.bus[idx].mes = clean(r['Semaine Prév. MES']);
        // Lot depuis HME (normalisation "LOT 2" → "LOT2")
        if (r['LOT']) {
          const lot = parseLot(r['LOT']);
          if (lot && lot !== 'NONDÉMARRÉ' && lot !== 'NONDEMARRÉ') db.bus[idx].lot = lot;
        }
      });
      saveTable('bus', db.bus);
    }

    // ── 3. Suivi TRAVAUX ─────────────────────────
    // Colonnes: LOT | Typologie | RR | Plans reçus | Date Réception | NB BPE alignées |
    //           Linéaire tiré | Date prév BUS | Date prév Pendulaires | Etat BUS |
    //           ... | Date livraison | ... | Points bloquants / Risques | Projection |
    //           Commentaires BE | Commentaires RR
    const trxSheet = wb.Sheets['Suivi TRAVAUX'];
    if (trxSheet) {
      const rows = XLSX.utils.sheet_to_json(trxSheet, { defval: null });
      result.sheets['Suivi TRAVAUX'] = rows.length;
      rows.forEach((r, idx) => {
        const topo = clean(r['Typologie']);
        if (!topo) return;
        const bidx = db.bus.findIndex(b => b.id === topo);
        if (bidx < 0) return;

        // RR (peut contenir plusieurs personnes "A ; B" — on garde tout)
        if (r['RR']) db.bus[bidx].rr = clean(r['RR']);

        // Etat depuis Suivi TRAVAUX — plus granulaire que Etat BUS
        if (r['Etat BUS']) {
          const etatTrx = parseEtat(r['Etat BUS']);
          // On prend le max entre les deux onglets
          if (etatTrx > db.bus[bidx].etat) db.bus[bidx].etat = etatTrx;
        }

        // Risques / points bloquants → risque
        const risques = clean(r['Points bloquants / Risques ide'] || r['Points bloquants / Risques']);
        if (risques && risques.length > 3) {
          db.bus[bidx].risque = 'Élevé';
        }

        // Commentaires
        const cmtBE  = clean(r['Commentaires BE']);
        const cmtRR  = clean(r['Commentaires RR']);
        const cmtAll = [cmtBE, cmtRR].filter(Boolean).join(' | ');
        if (cmtAll && !db.bus[bidx].comment) db.bus[bidx].comment = cmtAll;

        // Date livraison prévisionnelle
        if (r['Date livraison'] && !db.bus[bidx].liv) db.bus[bidx].liv = clean(r['Date livraison']);

        db.bus[bidx].updated_at = now();
      });
      saveTable('bus', db.bus);
    }

    // ── 4. SUIVI DELIVERY ────────────────────────
    // Ligne 1 : titres de groupes (SYNO INTER, HEPOC, CPM, Commande ferme)
    // Ligne 2 : vrais en-têtes (headers réels)
    // Colonnes clés:
    //   Col 0: LOT | Col 1: Projet (= topo) | Col 2: Zone | Col 5: Etat
    //   Col 6: HEPOC Transmis Orange | Col 7: HEPOC Envoyé à Orange
    //   Col 9: CPM Transmis Orange   | Col 10: CPM Envoyé à Orange
    //   Col 13: Commande ferme HEPOC | Col 14: Commande ferme CPM
    const dlvSheet = wb.Sheets['SUIVI DELIVERY'];
    if (dlvSheet) {
      // On saute la ligne 1 (titres groupes) et on utilise la ligne 2 comme header
      const rows = XLSX.utils.sheet_to_json(dlvSheet, { defval: null, range: 1 });
      result.sheets['SUIVI DELIVERY'] = rows.length;
      rows.forEach((r, idx) => {
        // La ligne 0 après range:1 est encore l'en-tête réel — skip si c'est "LOT"
        const topo = clean(r['Projet']);
        if (!topo || topo === 'Projet') return;
        const lot  = parseLot(r['LOT']);
        const row  = {
          topo, lot,
          etat:      clean(r['Etat'] || ''),
          hRcv:      parseOui(r['Transmis par Orange\nDate']),
          hF:        'NON', // pas de colonne dédiée dans cet Excel
          hRnv:      parseOui(r['Envoyé à Orange\nDate']),
          cRcv:      parseOui(r['Transmis par Orange\nDate_1'] || r['Transmis par Orange Date']),
          cF:        'NON',
          cRnv:      parseOui(r['Envoyé à Orange\nDate_1']    || r['Envoyé à Orange Date']),
          vO:        'NON',
          cf:        parseOui(r['HEPOC\nDate']) === 'OUI' && parseOui(r['CPM\nDate']) === 'OUI' ? 'OUI' : 'NON',
          updated_at: now(),
        };
        // Si l'état est "OK & Fini" → commande ferme OUI
        if (clean(r['Etat']).toLowerCase().includes('ok') || clean(r['Etat']).toLowerCase().includes('fini')) {
          row.cf = 'OUI'; row.vO = 'OUI';
        }
        const bidx = db.delivery.findIndex(d => d.topo === topo);
        if (bidx >= 0) {
          // Préserver les modifications manuelles si déjà dans la base
          db.delivery[bidx] = { ...db.delivery[bidx], ...row };
        } else {
          db.delivery.push(row);
        }
      });
      saveTable('delivery', db.delivery);
    }

    // ── 5. Mesures ───────────────────────────────
    // Colonnes: RR | Topo | Nb Tronçons | Accès Site MOE | Nb Tronçons Mesurés | Nb tronçons Ok | Nb tronçons Nok
    const mesSheet = wb.Sheets['Mesures'];
    if (mesSheet) {
      const rows = XLSX.utils.sheet_to_json(mesSheet, { defval: null });
      result.sheets['Mesures'] = rows.length;
      // Reset et réimport complet des mesures
      db.mesures = [];
      let mid = 1;
      rows.forEach(r => {
        const topo = clean(r['Topo']);
        if (!topo || topo === 'Topo') return;
        db.mesures.push({
          id:   mid++,
          topo,
          rr:   clean(r['RR']),
          nbT:  parseInt(r['Nb Tronçons'])  || 0,
          moe:  clean(r['Accès Site MOE'])  || 'Nok',
          nbM:  parseInt(r['Nb Tronçons Mesurés']) || 0,
          ok:   parseInt(r['Nb tronçons Ok'])  || 0,
          nok:  parseInt(r['Nb tronçons Nok']) || 0,
          updated_at: now(),
        });
      });
      saveTable('mesures', db.mesures);
    }

    // ── 6. CDD ───────────────────────────────────
    // Ligne 1 : avertissement "!!!!!!! MAJ PAR EQUIPE PROJET"
    // Ligne 2 : vrais en-têtes
    // Colonnes: LOT | Référence topologie | Département | RR | Site Code 42C | Nom site |
    //           Date MAD Orange | COL / SIN3 | Numèro Commande | COL prestation |
    //           Projet PFTO lancé | Info maj SIN3 | Commentaire Equipe PROJET
    const cddSheet = wb.Sheets['CDD'];
    if (cddSheet) {
      const rows = XLSX.utils.sheet_to_json(cddSheet, { defval: null, range: 1 });
      result.sheets['CDD'] = rows.length;
      db.cdd = [];
      let cid = 1;
      rows.forEach((r, idx) => {
        const topo = clean(r['Référence topologie']);
        if (!topo || topo === 'Référence topologie') return;
        const pftoRaw = r['Projet PFTO lancé par équipe o'] || r['Projet PFTO lancé'];
        const sin3Raw = r['Info maj sur SIN3 par Equipe p']  || r['Info maj SIN3'];
        db.cdd.push({
          id:      cid++,
          lot:     parseLot(r['LOT']),
          topo,
          dept:    clean(r['Département']),
          rr:      clean(r['RR']),
          sc:      clean(r['Site\nCode 42C'] || r['Site Code 42C']),
          ns:      clean(r['Nom site']),
          mad:     clean(r['Date MAD Orange \n'] || r['Date MAD Orange']),
          cmd:     clean(r['Numèro Commande']    || r['Numéro Commande']),
          pfto:    parseOui(pftoRaw),
          sin3:    parseOui(sin3Raw),
          comment: clean(r['Commentaire Equipe PROJET'] || ''),
          updated_at: now(),
        });
      });
      saveTable('cdd', db.cdd);
    }

    // ── 7. SUIVI EMPLACEMENT ─────────────────────
    // Ligne 1 : avertissement
    // Ligne 2 : en-têtes
    // Colonnes: Topo | Lot | Site | Dept | RR | Equipe opérationelle prévenue |
    //           Envoi prise de RDV | Annexe B signée | Installation des équipements |
    //           Consuel | Annexe C
    const emplSheet = wb.Sheets['SUIVI EMPLACEMENT'];
    if (emplSheet) {
      const rows = XLSX.utils.sheet_to_json(emplSheet, { defval: null, range: 1 });
      result.sheets['SUIVI EMPLACEMENT'] = rows.length;
      db.emplacements = [];
      let eid = 1;
      rows.forEach(r => {
        const topo = clean(r['Topo']);
        if (!topo || topo === 'Topo') return;
        db.emplacements.push({
          id:   eid++,
          topo,
          lot:  parseLot(r['Lot']),
          site: clean(r['Site']),
          dept: clean(r['Dept']),
          rr:   clean(r['RR']),
          ep:   parseOui(r['Equipe opérationelle prévenue']),
          rdv:  parseOui(r['Envoi prise de RDV ']),
          aB:   parseOui(r['Annexe B signée']),
          inst: parseOui(r['Installation des équipements']),
          cons: parseOui(r['Consuel']),
          aC:   parseOui(r['Annexe C']),
          updated_at: now(),
        });
      });
      saveTable('emplacements', db.emplacements);
    }

    result.warnings = result.warnings.slice(0, 10);
    res.json(result);

  } catch(e) {
    console.error('Import error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── HEALTH & SPA ──────────────────────────────────
app.get('/health', (req, res) => res.json({ status:'ok', time:now() }));
app.get('*', (req, res) => res.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅  NetCollect → http://localhost:${PORT}`);
  console.log(`📁  Données   → ${DATA_DIR}\n`);
});
