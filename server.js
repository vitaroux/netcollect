const express       = require('express');
const session       = require('express-session');
const bcrypt        = require('bcryptjs');
const multer        = require('multer');
const XLSX          = require('xlsx');
const cors          = require('cors');
const path          = require('path');
const fs            = require('fs');

const app  = express();
const PORT = process.env.PORT || 3000;
const USE_PG = !!process.env.DATABASE_URL;
const SESSION_SECRET = process.env.SESSION_SECRET || 'nc_secret_change_in_prod_' + Date.now();

// ── POSTGRESQL ────────────────────────────────────
let pool = null;
if (USE_PG) {
  const { Pool } = require('pg');
  pool = new Pool({ connectionString: process.env.DATABASE_URL, ssl:{ rejectUnauthorized:false } });
  console.log('🐘 Mode PostgreSQL');
} else {
  console.log('📁 Mode fichiers JSON');
}

// ── JSON STORAGE ──────────────────────────────────
const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function loadJ(name) {
  const f = path.join(DATA_DIR, name+'.json');
  try { return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f,'utf8')) : []; } catch(e){ return []; }
}
function saveJ(name, data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR,{recursive:true});
  fs.writeFileSync(path.join(DATA_DIR,name+'.json'), JSON.stringify(data,null,2));
}

// ── HELPERS ───────────────────────────────────────
const now      = () => new Date().toISOString();
const clean    = v => v===null||v===undefined?'':String(v).replace(/\xa0/g,' ').trim();
const parseLot = v => v?String(v).trim().toUpperCase().replace(/\s+/g,''):'';
const parseOui = v => v&&String(v).trim().toUpperCase().startsWith('OUI')?'OUI':'NON';
function parseEtat(val) {
  if (!val) return 0;
  const s = String(val).trim().toLowerCase();
  if (s.startsWith('0')||s.includes('non dém')||s.includes('non dem')) return 0;
  if (s.startsWith('1')||s.includes('design'))   return 1;
  if (s.startsWith('2')||s.includes('delivery')) return 2;
  if (s.startsWith('3')||s.includes('plan')||s.includes('étude')||s.includes('analyse')) return 3;
  if (s.startsWith('4')||s.includes('travaux'))  return 4;
  if (s.startsWith('5')||s==='livré'||s.includes('livr')) return 5;
  if (s.startsWith('6')||s.includes('mes'))      return 6;
  return 0;
}
async function logH(entity, eid, field, oldV, newV, user='système') {
  if (String(oldV??'') === String(newV??'')) return;
  if (USE_PG) {
    await pool.query('INSERT INTO history(entity,entity_id,field,old_val,new_val,user_name) VALUES($1,$2,$3,$4,$5,$6)',
      [entity,eid,field,String(oldV??''),String(newV??''),user]);
  } else {
    const h=loadJ('history');
    h.push({id:Date.now()+Math.random(),entity,entity_id:eid,field,
      old_val:String(oldV??''),new_val:String(newV??''),user_name:user,created_at:now()});
    saveJ('history',h);
  }
}

const PROFILES = ['Admin','CDP','CDR','BE','Projet'];

// ── SESSION STORE ─────────────────────────────────
let sessionStore;
if (USE_PG) {
  try {
    const PgSession = require('connect-pg-simple')(session);
    sessionStore = new PgSession({ pool, createTableIfMissing: true });
  } catch(e) { console.log('connect-pg-simple non dispo, sessions en mémoire'); }
}

// ── MIDDLEWARE ────────────────────────────────────
app.use(cors({ origin: true, credentials: true }));
app.use(express.json());
app.use(session({
  store: sessionStore,
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: false,      // true si HTTPS en prod
    httpOnly: true,
    maxAge: 8 * 60 * 60 * 1000  // 8h
  }
}));
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50*1024*1024 } });

// ── AUTH MIDDLEWARE ───────────────────────────────
function requireAuth(req, res, next) {
  if (req.session?.userId) return next();
  res.status(401).json({ error: 'Non authentifié', redirect: '/login' });
}

// ── DB SCHEMA INIT ────────────────────────────────
async function initDB() {
  if (!USE_PG) return;
  await pool.query(`
    CREATE TABLE IF NOT EXISTS bus (
      id TEXT PRIMARY KEY, lot TEXT DEFAULT '', dor TEXT DEFAULT '',
      ext TEXT DEFAULT '', etat INTEGER DEFAULT 0, rr TEXT DEFAULT '',
      av INTEGER DEFAULT 0, liv TEXT DEFAULT '', mes TEXT DEFAULT '',
      risque TEXT DEFAULT 'Faible', comment TEXT DEFAULT '',
      created_at TIMESTAMPTZ DEFAULT NOW(), updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS delivery (
      topo TEXT PRIMARY KEY, lot TEXT DEFAULT '', etat TEXT DEFAULT '',
      h_rcv TEXT DEFAULT 'NON', h_f TEXT DEFAULT 'NON', h_rnv TEXT DEFAULT 'NON',
      c_rcv TEXT DEFAULT 'NON', c_f TEXT DEFAULT 'NON', c_rnv TEXT DEFAULT 'NON',
      v_o TEXT DEFAULT 'NON', cf TEXT DEFAULT 'NON', updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS mesures (
      id SERIAL PRIMARY KEY, topo TEXT, rr TEXT DEFAULT '', nb_t INTEGER DEFAULT 0,
      moe TEXT DEFAULT 'Nok', nb_m INTEGER DEFAULT 0, ok INTEGER DEFAULT 0,
      nok INTEGER DEFAULT 0, updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS cdd (
      id SERIAL PRIMARY KEY, lot TEXT DEFAULT '', topo TEXT DEFAULT '',
      dept TEXT DEFAULT '', rr TEXT DEFAULT '', sc TEXT DEFAULT '', ns TEXT DEFAULT '',
      mad TEXT DEFAULT '', cmd TEXT DEFAULT '', pfto TEXT DEFAULT 'NON',
      sin3 TEXT DEFAULT 'NON', comment TEXT DEFAULT '', updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS emplacements (
      id SERIAL PRIMARY KEY, topo TEXT DEFAULT '', lot TEXT DEFAULT '',
      site TEXT DEFAULT '', dept TEXT DEFAULT '', rr TEXT DEFAULT '',
      ep TEXT DEFAULT 'NON', rdv TEXT DEFAULT 'NON', a_b TEXT DEFAULT 'NON',
      inst TEXT DEFAULT 'NON', cons TEXT DEFAULT 'NON', a_c TEXT DEFAULT 'NON',
      updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS cpm (
      topo TEXT PRIMARY KEY, ref_client TEXT DEFAULT '', ref_commande TEXT DEFAULT '',
      date_crmad TEXT DEFAULT '', statut TEXT DEFAULT 'Non reçu',
      commentaire TEXT DEFAULT '', date_reception TEXT DEFAULT '', anomalies TEXT DEFAULT '',
      nb_segments INTEGER DEFAULT 0, nb_sites INTEGER DEFAULT 0,
      nb_fo_total INTEGER DEFAULT 0, longueur_totale INTEGER DEFAULT 0,
      segments JSONB DEFAULT '[]', sites JSONB DEFAULT '[]', hepoc JSONB DEFAULT '[]',
      lom JSONB DEFAULT '[]', loir JSONB DEFAULT '[]', penetration JSONB DEFAULT '[]',
      imported_at TIMESTAMPTZ DEFAULT NOW(), updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY, nom TEXT NOT NULL, prenom TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE, profil TEXT NOT NULL DEFAULT 'Projet',
      password_hash TEXT DEFAULT NULL,
      actif BOOLEAN DEFAULT TRUE,
      created_at TIMESTAMPTZ DEFAULT NOW(), updated_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS history (
      id SERIAL PRIMARY KEY, entity TEXT, entity_id TEXT, field TEXT,
      old_val TEXT, new_val TEXT, user_name TEXT DEFAULT 'système',
      created_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS comments (
      id SERIAL PRIMARY KEY, topo TEXT, author TEXT DEFAULT 'Équipe',
      body TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE INDEX IF NOT EXISTS idx_history_eid ON history(entity_id);
    CREATE INDEX IF NOT EXISTS idx_comments_topo ON comments(topo);
  `);
  // Ajouter password_hash si absent (migration)
  await pool.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS password_hash TEXT DEFAULT NULL`).catch(()=>{});
  await pool.query(`ALTER TABLE cpm ADD COLUMN IF NOT EXISTS lom JSONB DEFAULT '[]'`).catch(()=>{});
  await pool.query(`ALTER TABLE cpm ADD COLUMN IF NOT EXISTS loir JSONB DEFAULT '[]'`).catch(()=>{});
  await pool.query(`ALTER TABLE cpm ADD COLUMN IF NOT EXISTS penetration JSONB DEFAULT '[]'`).catch(()=>{});
  console.log('✅ Schéma PostgreSQL initialisé');
}

async function initUsers() {
  const defaultHash = await bcrypt.hash('Admin123!', 10);
  if (USE_PG) {
    const {rows} = await pool.query('SELECT COUNT(*)::int as n FROM users');
    if (rows[0].n === 0) {
      await pool.query(
        'INSERT INTO users(nom,prenom,email,profil,password_hash) VALUES($1,$2,$3,$4,$5) ON CONFLICT(email) DO NOTHING',
        ['Admin','TransProd','admin@TransProd.fr','Admin',defaultHash]
      );
      console.log('👤 Admin créé : admin@TransProd.fr / Admin123!');
    }
    // Mettre un hash par défaut aux utilisateurs sans mot de passe
    await pool.query(`UPDATE users SET password_hash=$1 WHERE password_hash IS NULL OR password_hash=''`,[defaultHash]);
  } else {
    const users = loadJ('users');
    if (users.length === 0) {
      saveJ('users', [{id:1,nom:'Admin',prenom:'TransProd',email:'admin@TransProd.fr',
        profil:'Admin',password_hash:defaultHash,actif:true,created_at:now(),updated_at:now()}]);
      console.log('👤 Admin créé : admin@TransProd.fr / Admin123!');
    } else {
      let changed = false;
      users.forEach(u => { if (!u.password_hash) { u.password_hash = defaultHash; changed = true; } });
      if (changed) saveJ('users', users);
    }
  }
}

// ═══════════════════════════════════════════════
// AUTH ROUTES (publiques)
// ═══════════════════════════════════════════════

// Vérifier la session courante
app.get('/api/auth/me', (req, res) => {
  if (!req.session?.userId) return res.status(401).json({ authenticated: false });
  res.json({ authenticated: true, user: req.session.user });
});

// Login
app.post('/api/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ error: 'Email et mot de passe requis' });

    let user;
    if (USE_PG) {
      const {rows} = await pool.query('SELECT * FROM users WHERE LOWER(email)=LOWER($1) AND actif=TRUE',[email.trim()]);
      user = rows[0];
    } else {
      const users = loadJ('users');
      user = users.find(u => u.email.toLowerCase() === email.trim().toLowerCase() && u.actif !== false);
    }

    if (!user) return res.status(401).json({ error: 'Compte introuvable' });
    if (!user.profil) return res.status(403).json({ error: 'Aucun profil affecté. Contactez un administrateur.' });
    if (!user.password_hash) return res.status(401).json({ error: 'Mot de passe non défini. Contactez un administrateur.' });

    const ok = await bcrypt.compare(password, user.password_hash);
    if (!ok) return res.status(401).json({ error: 'Mot de passe incorrect' });

    // Créer la session
    const sessionUser = { id: user.id, nom: user.nom, prenom: user.prenom, email: user.email, profil: user.profil };
    req.session.userId = user.id;
    req.session.user   = sessionUser;

    res.json({ ok: true, user: sessionUser });
  } catch(e) { console.error('Login error:', e); res.status(500).json({ error: e.message }); }
});

// Logout
app.post('/api/auth/logout', (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// ── ROUTES PROTÉGÉES : tout ce qui suit nécessite requireAuth ──

// ═══════════════════════════════════════════════
// USERS / PROFILS
// ═══════════════════════════════════════════════
app.get('/api/users', requireAuth, async (req,res) => {
  try {
    if (USE_PG) {
      const {rows} = await pool.query('SELECT id,nom,prenom,email,profil,actif,created_at,updated_at FROM users ORDER BY nom,prenom');
      return res.json(rows);
    }
    res.json(loadJ('users').map(u=>({...u,password_hash:undefined})).sort((a,b)=>a.nom.localeCompare(b.nom)));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.post('/api/users', requireAuth, async (req,res) => {
  try {
    const {nom,prenom,email,profil,password} = req.body;
    if (!nom?.trim())   return res.status(400).json({error:'Le nom est obligatoire'});
    if (!prenom?.trim())return res.status(400).json({error:'Le prénom est obligatoire'});
    if (!email?.trim()) return res.status(400).json({error:"L'adresse mail est obligatoire"});
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) return res.status(400).json({error:'Adresse mail invalide'});
    if (!profil||!PROFILES.includes(profil)) return res.status(400).json({error:'Profil invalide'});

    const pwd = password?.trim() || 'Temp1234!';
    const hash = await bcrypt.hash(pwd, 10);

    if (USE_PG) {
      try {
        const {rows} = await pool.query(
          'INSERT INTO users(nom,prenom,email,profil,password_hash) VALUES($1,$2,$3,$4,$5) RETURNING id,nom,prenom,email,profil,actif,created_at',
          [nom.trim(),prenom.trim(),email.trim().toLowerCase(),profil,hash]
        );
        return res.status(201).json(rows[0]);
      } catch(e) {
        if (e.code==='23505') return res.status(400).json({error:'Email déjà utilisée'});
        throw e;
      }
    }
    const db = loadJ('users');
    if (db.find(u=>u.email===email.trim().toLowerCase())) return res.status(400).json({error:'Email déjà utilisée'});
    const row = {id:Date.now(),nom:nom.trim(),prenom:prenom.trim(),email:email.trim().toLowerCase(),
      profil,password_hash:hash,actif:true,created_at:now(),updated_at:now()};
    db.push(row); saveJ('users',db);
    const {password_hash,...safe}=row;
    res.status(201).json(safe);
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.put('/api/users/:id', requireAuth, async (req,res) => {
  try {
    const {nom,prenom,email,profil,actif,password} = req.body;
    if (nom!==undefined&&!nom?.trim())     return res.status(400).json({error:'Nom obligatoire'});
    if (prenom!==undefined&&!prenom?.trim())return res.status(400).json({error:'Prénom obligatoire'});
    if (email!==undefined) {
      if (!email?.trim()) return res.status(400).json({error:'Email obligatoire'});
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) return res.status(400).json({error:'Email invalide'});
    }
    if (profil!==undefined&&!PROFILES.includes(profil)) return res.status(400).json({error:'Profil invalide'});

    let hash = null;
    if (password?.trim()) hash = await bcrypt.hash(password.trim(), 10);

    if (USE_PG) {
      const fields=[]; const vals=[];
      if(nom    !==undefined){fields.push(`nom=$${fields.length+1}`);    vals.push(nom.trim());}
      if(prenom !==undefined){fields.push(`prenom=$${fields.length+1}`); vals.push(prenom.trim());}
      if(email  !==undefined){fields.push(`email=$${fields.length+1}`);  vals.push(email.trim().toLowerCase());}
      if(profil !==undefined){fields.push(`profil=$${fields.length+1}`); vals.push(profil);}
      if(actif  !==undefined){fields.push(`actif=$${fields.length+1}`);  vals.push(actif);}
      if(hash){fields.push(`password_hash=$${fields.length+1}`); vals.push(hash);}
      if(!fields.length) return res.status(400).json({error:'Aucun champ'});
      try {
        const {rows} = await pool.query(
          `UPDATE users SET ${fields.join(',')},updated_at=NOW() WHERE id=$${fields.length+1} RETURNING id,nom,prenom,email,profil,actif,updated_at`,
          [...vals,req.params.id]
        );
        if (!rows[0]) return res.status(404).json({error:'Utilisateur introuvable'});
        return res.json(rows[0]);
      } catch(e) {
        if (e.code==='23505') return res.status(400).json({error:'Email déjà utilisée'});
        throw e;
      }
    }
    const db=loadJ('users'); const idx=db.findIndex(u=>u.id===+req.params.id);
    if(idx<0) return res.status(404).json({error:'Utilisateur introuvable'});
    if(nom    !==undefined) db[idx].nom    = nom.trim();
    if(prenom !==undefined) db[idx].prenom = prenom.trim();
    if(email  !==undefined) {
      if(db.find((u,i)=>u.email===email.trim().toLowerCase()&&i!==idx)) return res.status(400).json({error:'Email déjà utilisée'});
      db[idx].email = email.trim().toLowerCase();
    }
    if(profil !==undefined) db[idx].profil = profil;
    if(actif  !==undefined) db[idx].actif  = actif;
    if(hash) db[idx].password_hash = hash;
    db[idx].updated_at=now(); saveJ('users',db);
    const {password_hash,...safe}=db[idx];
    res.json(safe);
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.delete('/api/users/:id', requireAuth, async (req,res) => {
  try {
    // Empêcher la suppression de son propre compte
    if (String(req.session.userId) === String(req.params.id))
      return res.status(400).json({error:'Vous ne pouvez pas supprimer votre propre compte'});
    if (USE_PG) {
      const {rows} = await pool.query('DELETE FROM users WHERE id=$1 RETURNING id',[req.params.id]);
      if (!rows[0]) return res.status(404).json({error:'Utilisateur introuvable'});
      return res.json({ok:true});
    }
    const db=loadJ('users'); const idx=db.findIndex(u=>u.id===+req.params.id);
    if(idx<0) return res.status(404).json({error:'Utilisateur introuvable'});
    db.splice(idx,1); saveJ('users',db); res.json({ok:true});
  } catch(e){ res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// BUS
// ═══════════════════════════════════════════════
app.get('/api/bus', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM bus ORDER BY lot,id'); return res.json(rows); }
    res.json(loadJ('bus').sort((a,b)=>(a.lot+a.id).localeCompare(b.lot+b.id)));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.get('/api/bus/:id', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM bus WHERE id=$1',[req.params.id]); if(!rows[0]) return res.status(404).json({error:'BUS introuvable'}); return res.json(rows[0]); }
    const row=loadJ('bus').find(b=>b.id===req.params.id);
    if(!row) return res.status(404).json({error:'BUS introuvable'});
    res.json(row);
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.put('/api/bus/:id', requireAuth, async (req,res) => {
  try {
    const {id}=req.params; const user=req.body.user||req.session.user?.prenom||'système';
    const allowed=['lot','dor','ext','etat','rr','av','liv','mes','risque','comment'];
    const fields=allowed.filter(k=>req.body[k]!==undefined);
    if (!fields.length) return res.status(400).json({error:'Aucun champ'});
    if (USE_PG) {
      const {rows:cur}=await pool.query('SELECT * FROM bus WHERE id=$1',[id]);
      if (!cur[0]) return res.status(404).json({error:'BUS introuvable'});
      for (const k of fields) await logH('bus',id,k,cur[0][k],req.body[k],user);
      const sets=fields.map((k,i)=>`${k}=$${i+1}`).join(',');
      const {rows}=await pool.query(`UPDATE bus SET ${sets},updated_at=NOW() WHERE id=$${fields.length+1} RETURNING *`,[...fields.map(k=>req.body[k]),id]);
      return res.json(rows[0]);
    }
    const db=loadJ('bus'); const idx=db.findIndex(b=>b.id===id);
    if (idx<0) return res.status(404).json({error:'BUS introuvable'});
    for (const k of fields) await logH('bus',id,k,db[idx][k],req.body[k],user);
    fields.forEach(k=>db[idx][k]=req.body[k]); db[idx].updated_at=now();
    saveJ('bus',db); res.json(db[idx]);
  } catch(e){ res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// DELIVERY
// ═══════════════════════════════════════════════
app.get('/api/delivery', requireAuth, async (req,res) => {
  try {
    if (USE_PG) {
      const {rows}=await pool.query('SELECT * FROM delivery ORDER BY lot,topo');
      return res.json(rows.map(r=>({topo:r.topo,lot:r.lot,etat:r.etat,hRcv:r.h_rcv,hF:r.h_f,hRnv:r.h_rnv,cRcv:r.c_rcv,cF:r.c_f,cRnv:r.c_rnv,vO:r.v_o,cf:r.cf})));
    }
    res.json(loadJ('delivery'));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.put('/api/delivery/:topo', requireAuth, async (req,res) => {
  try {
    const f=req.body; const t=req.params.topo;
    if (USE_PG) {
      await pool.query(`INSERT INTO delivery(topo,lot,etat,h_rcv,h_f,h_rnv,c_rcv,c_f,c_rnv,v_o,cf) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
        ON CONFLICT(topo) DO UPDATE SET lot=$2,etat=$3,h_rcv=$4,h_f=$5,h_rnv=$6,c_rcv=$7,c_f=$8,c_rnv=$9,v_o=$10,cf=$11,updated_at=NOW()`,
        [t,f.lot||'',f.etat||'',f.hRcv||'NON',f.hF||'NON',f.hRnv||'NON',f.cRcv||'NON',f.cF||'NON',f.cRnv||'NON',f.vO||'NON',f.cf||'NON']);
      return res.json({ok:true});
    }
    const db=loadJ('delivery'); const idx=db.findIndex(d=>d.topo===t);
    const row={topo:t,lot:f.lot||'',etat:f.etat||'',hRcv:f.hRcv||'NON',hF:f.hF||'NON',hRnv:f.hRnv||'NON',cRcv:f.cRcv||'NON',cF:f.cF||'NON',cRnv:f.cRnv||'NON',vO:f.vO||'NON',cf:f.cf||'NON',updated_at:now()};
    if(idx>=0)db[idx]=row;else db.push(row);
    saveJ('delivery',db); res.json({ok:true});
  } catch(e){ res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// MESURES, CDD, EMPLACEMENTS, CPM
// ═══════════════════════════════════════════════
app.get('/api/mesures', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM mesures ORDER BY topo'); return res.json(rows.map(r=>({...r,nbT:r.nb_t,nbM:r.nb_m}))); }
    res.json(loadJ('mesures'));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.get('/api/cdd', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM cdd ORDER BY lot,topo'); return res.json(rows); }
    res.json(loadJ('cdd'));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.put('/api/cdd/:id', requireAuth, async (req,res) => {
  try {
    const f=req.body; const user=req.session.user?.prenom||'système';
    if (USE_PG) {
      const {rows:cur}=await pool.query('SELECT * FROM cdd WHERE id=$1',[req.params.id]);
      if (!cur[0]) return res.status(404).json({error:'CDD introuvable'});
      for (const k of ['pfto','sin3','comment']) if(f[k]!==undefined) await logH('cdd',req.params.id,k,cur[0][k],f[k],user);
      await pool.query('UPDATE cdd SET pfto=$1,sin3=$2,comment=$3,updated_at=NOW() WHERE id=$4',[f.pfto||'NON',f.sin3||'NON',f.comment||'',req.params.id]);
      return res.json({ok:true});
    }
    const db=loadJ('cdd'); const idx=db.findIndex(c=>c.id===+req.params.id);
    if(idx<0) return res.status(404).json({error:'CDD introuvable'});
    for (const k of ['pfto','sin3','comment']) if(f[k]!==undefined){await logH('cdd',req.params.id,k,db[idx][k],f[k],user);db[idx][k]=f[k];}
    db[idx].updated_at=now(); saveJ('cdd',db); res.json({ok:true});
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.get('/api/emplacements', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM emplacements ORDER BY lot,topo'); return res.json(rows.map(r=>({...r,aB:r.a_b,aC:r.a_c}))); }
    res.json(loadJ('emplacements'));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.get('/api/cpm', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM cpm ORDER BY topo'); return res.json(rows); }
    res.json(loadJ('cpm'));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.put('/api/cpm/:topo', requireAuth, async (req,res) => {
  try {
    const f=req.body; const user=req.session.user?.prenom||'système'; const t=req.params.topo;
    if (USE_PG) {
      const {rows:cur}=await pool.query('SELECT * FROM cpm WHERE topo=$1',[t]);
      if (!cur[0]) return res.status(404).json({error:'CPM introuvable'});
      for (const k of ['statut','commentaire','date_reception','anomalies']) if(f[k]!==undefined) await logH('cpm',t,k,cur[0][k],f[k],user);
      await pool.query('UPDATE cpm SET statut=$1,commentaire=$2,date_reception=$3,anomalies=$4,updated_at=NOW() WHERE topo=$5',[f.statut||'Non reçu',f.commentaire||'',f.date_reception||'',f.anomalies||'',t]);
      const {rows}=await pool.query('SELECT * FROM cpm WHERE topo=$1',[t]); return res.json(rows[0]);
    }
    const db=loadJ('cpm'); const idx=db.findIndex(c=>c.topo===t);
    if(idx<0) return res.status(404).json({error:'CPM introuvable'});
    for (const k of ['statut','commentaire','date_reception','anomalies']) if(f[k]!==undefined){await logH('cpm',t,k,db[idx][k],f[k],user);db[idx][k]=f[k];}
    db[idx].updated_at=now(); saveJ('cpm',db); res.json(db[idx]);
  } catch(e){ res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// HISTORY & COMMENTS
// ═══════════════════════════════════════════════
app.get('/api/history/:id', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM history WHERE entity_id=$1 ORDER BY created_at DESC LIMIT 100',[req.params.id]); return res.json(rows); }
    res.json(loadJ('history').filter(r=>r.entity_id===req.params.id).sort((a,b)=>b.created_at.localeCompare(a.created_at)).slice(0,100));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.get('/api/comments/:topo', requireAuth, async (req,res) => {
  try {
    if (USE_PG) { const {rows}=await pool.query('SELECT * FROM comments WHERE topo=$1 ORDER BY created_at DESC',[req.params.topo]); return res.json(rows); }
    res.json(loadJ('comments').filter(r=>r.topo===req.params.topo).sort((a,b)=>b.created_at.localeCompare(a.created_at)));
  } catch(e){ res.status(500).json({error:e.message}); }
});

app.post('/api/comments/:topo', requireAuth, async (req,res) => {
  try {
    const {body:text}=req.body;
    const author = req.session.user ? req.session.user.prenom+' '+req.session.user.nom : 'Équipe';
    if (!text?.trim()) return res.status(400).json({error:'Commentaire vide'});
    if (USE_PG) { const {rows}=await pool.query('INSERT INTO comments(topo,author,body) VALUES($1,$2,$3) RETURNING *',[req.params.topo,author,text.trim()]); return res.status(201).json(rows[0]); }
    const db=loadJ('comments'); const row={id:Date.now(),topo:req.params.topo,author,body:text.trim(),created_at:now()};
    db.push(row); saveJ('comments',db); res.status(201).json(row);
  } catch(e){ res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// STATS
// ═══════════════════════════════════════════════
app.get('/api/stats', requireAuth, async (req,res) => {
  try {
    if (USE_PG) {
      const [e1,e2,e3,e4,e5,e6,e7]=await Promise.all([
        pool.query('SELECT etat,COUNT(*)::int as count FROM bus GROUP BY etat'),
        pool.query(`SELECT lot,COUNT(*)::int as total,SUM(CASE WHEN etat=4 THEN 1 ELSE 0 END)::int as trx,SUM(CASE WHEN etat>=5 THEN 1 ELSE 0 END)::int as livres,ROUND(AVG(av))::int as "avMoy" FROM bus GROUP BY lot ORDER BY lot`),
        pool.query('SELECT ROUND(AVG(av))::int as avg FROM bus'),
        pool.query("SELECT COUNT(*)::int as n FROM delivery WHERE cf='OUI'"),
        pool.query('SELECT COALESCE(SUM(nok),0)::int as n FROM mesures'),
        pool.query("SELECT COUNT(*)::int as n FROM cdd WHERE sin3='NON'"),
        pool.query("SELECT COUNT(*)::int as n FROM bus WHERE risque='Élevé'"),
      ]);
      return res.json({byEtat:e1.rows,byLot:e2.rows,avGlobal:e3.rows[0]?.avg||0,cfCount:e4.rows[0]?.n||0,nokCount:e5.rows[0]?.n||0,sin3Nok:e6.rows[0]?.n||0,riskHigh:e7.rows[0]?.n||0});
    }
    const bus=loadJ('bus'),mes=loadJ('mesures'),cdd=loadJ('cdd'),dlv=loadJ('delivery');
    const byEtat=[0,1,2,3,4,5,6].map(e=>({etat:e,count:bus.filter(b=>b.etat===e).length})).filter(r=>r.count>0);
    const lots=[...new Set(bus.map(b=>b.lot).filter(Boolean))].sort();
    const byLot=lots.map(lot=>{const lb=bus.filter(b=>b.lot===lot);return{lot,total:lb.length,trx:lb.filter(b=>b.etat===4).length,livres:lb.filter(b=>b.etat>=5).length,avMoy:lb.length?Math.round(lb.reduce((a,b)=>a+b.av,0)/lb.length):0};});
    res.json({byEtat,byLot,avGlobal:bus.length?Math.round(bus.reduce((a,b)=>a+b.av,0)/bus.length):0,cfCount:dlv.filter(d=>d.cf==='OUI').length,nokCount:mes.reduce((a,m)=>a+(m.nok||0),0),sin3Nok:cdd.filter(c=>c.sin3==='NON').length,riskHigh:bus.filter(b=>b.risque==='Élevé').length});
  } catch(e){ res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// IMPORT EXCEL
// ═══════════════════════════════════════════════
app.post('/api/import', requireAuth, upload.single('file'), async (req,res) => {
  if (!req.file) return res.status(400).json({error:'Aucun fichier reçu'});
  try {
    const wb=XLSX.read(req.file.buffer,{type:'buffer'});
    const result={imported:0,updated:0,sheets:{},warnings:[],errors:[]};
    async function upsertBus(topo,data) {
      if (USE_PG) {
        const ex=await pool.query('SELECT id FROM bus WHERE id=$1',[topo]);
        await pool.query(`INSERT INTO bus(id,lot,dor,ext,etat) VALUES($1,$2,$3,$4,$5) ON CONFLICT(id) DO UPDATE SET lot=EXCLUDED.lot,dor=EXCLUDED.dor,ext=EXCLUDED.ext,etat=EXCLUDED.etat,updated_at=NOW()`,[topo,data.lot||'',data.dor||'',data.ext||'',data.etat??0]);
        if(ex.rows.length)result.updated++;else result.imported++;
      } else {
        const db=loadJ('bus'); const idx=db.findIndex(b=>b.id===topo);
        if(idx>=0){Object.assign(db[idx],{...data,updated_at:now()});result.updated++;}
        else{db.push({id:topo,...data,rr:'',av:0,liv:'',mes:'',risque:'Faible',comment:'',created_at:now(),updated_at:now()});result.imported++;}
        saveJ('bus',db);
      }
    }
    async function ubf(topo,field,val){
      if(USE_PG)await pool.query(`UPDATE bus SET ${field}=$1 WHERE id=$2`,[val,topo]);
      else{const db=loadJ('bus');const idx=db.findIndex(b=>b.id===topo);if(idx>=0){db[idx][field]=val;saveJ('bus',db);}}
    }
    async function ubMax(topo,field,val){
      if(USE_PG)await pool.query(`UPDATE bus SET ${field}=GREATEST(${field},$1) WHERE id=$2`,[val,topo]);
      else{const db=loadJ('bus');const idx=db.findIndex(b=>b.id===topo);if(idx>=0&&val>db[idx][field]){db[idx][field]=val;saveJ('bus',db);}}
    }
    const bsSheet=wb.Sheets['Etat BUS'];
    if(bsSheet){const rows=XLSX.utils.sheet_to_json(bsSheet,{defval:null});result.sheets['Etat BUS']=rows.length;for(const[idx,row]of rows.entries()){const topo=clean(row['Nom topologie']);if(!topo){result.warnings.push(`L${idx+2}: topologie vide`);continue;}await upsertBus(topo,{lot:parseLot(row['Lot']),dor:clean(row['DOR']),ext:clean(row['Extrémité']||row['Extrémités']),etat:parseEtat(row['Etat BUS'])});}}else result.errors.push("Onglet 'Etat BUS' introuvable");
    const hme=wb.Sheets['HME'];if(hme){const rows=XLSX.utils.sheet_to_json(hme,{defval:null});result.sheets['HME']=rows.length;for(const r of rows){const topo=clean(r['ReferenceRegroupement']);if(!topo)continue;const lot=parseLot(r['LOT']);if(lot&&!['NONDÉMARRÉ','NONDEMARRÉ',''].includes(lot))await ubf(topo,'lot',lot);if(r['Semaine Prév. Livraison'])await ubf(topo,'liv',clean(String(r['Semaine Prév. Livraison'])));if(r['Semaine Prév. MES'])await ubf(topo,'mes',clean(String(r['Semaine Prév. MES'])));}}
    const trx=wb.Sheets['Suivi TRAVAUX'];if(trx){const rows=XLSX.utils.sheet_to_json(trx,{defval:null});result.sheets['Suivi TRAVAUX']=rows.length;for(const r of rows){const topo=clean(r['Typologie']||r['Topologie']);if(!topo)continue;if(r['RR'])await ubf(topo,'rr',clean(r['RR']));const etatTrx=parseEtat(r['Etat BUS']);if(etatTrx>0)await ubMax(topo,'etat',etatTrx);const risques=clean(r['Points bloquants / Risques ide']||r['Points bloquants / Risques']||'');if(risques.length>3)await ubf(topo,'risque','Élevé');const cmt=[clean(r['Commentaires BE']),clean(r['Commentaires RR'])].filter(Boolean).join(' | ');if(cmt){if(USE_PG)await pool.query("UPDATE bus SET comment=$1 WHERE id=$2 AND (comment='' OR comment IS NULL)",[cmt,topo]);else{const db=loadJ('bus');const idx=db.findIndex(b=>b.id===topo);if(idx>=0&&!db[idx].comment){db[idx].comment=cmt;saveJ('bus',db);}}}}}
    const dlvSheet=wb.Sheets['SUIVI DELIVERY'];if(dlvSheet){const rows=XLSX.utils.sheet_to_json(dlvSheet,{defval:null,range:1});result.sheets['SUIVI DELIVERY']=rows.length;const dlvData=[];for(const r of rows){const topo=clean(r['Projet']);if(!topo||topo==='Projet')continue;const etat=clean(r['Etat']||'');const cf=etat.toLowerCase().includes('ok')||etat.toLowerCase().includes('fini')?'OUI':'NON';const row={topo,lot:parseLot(r['LOT']),etat,hRcv:parseOui(r['Transmis par Orange\nDate']),hF:'NON',hRnv:parseOui(r['Envoyé à Orange\nDate']),cRcv:parseOui(r['Transmis par Orange\nDate_1']||r['Transmis par Orange Date']),cF:'NON',cRnv:parseOui(r['Envoyé à Orange\nDate_1']||r['Envoyé à Orange Date']),vO:cf==='OUI'?'OUI':'NON',cf};if(USE_PG)await pool.query(`INSERT INTO delivery(topo,lot,etat,h_rcv,h_f,h_rnv,c_rcv,c_f,c_rnv,v_o,cf) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) ON CONFLICT(topo) DO UPDATE SET lot=$2,etat=$3,h_rcv=$4,h_f=$5,h_rnv=$6,c_rcv=$7,c_f=$8,c_rnv=$9,v_o=$10,cf=$11,updated_at=NOW()`,[row.topo,row.lot,row.etat,row.hRcv,row.hF,row.hRnv,row.cRcv,row.cF,row.cRnv,row.vO,row.cf]);else dlvData.push(row);}if(!USE_PG)saveJ('delivery',dlvData);}
    const mesSheet=wb.Sheets['Mesures'];if(mesSheet){const rows=XLSX.utils.sheet_to_json(mesSheet,{defval:null});result.sheets['Mesures']=rows.length;if(USE_PG)await pool.query('DELETE FROM mesures');const mesData=[];let mid=1;for(const r of rows){const topo=clean(r['Topo']);if(!topo||topo==='Topo')continue;if(USE_PG)await pool.query('INSERT INTO mesures(topo,rr,nb_t,moe,nb_m,ok,nok) VALUES($1,$2,$3,$4,$5,$6,$7)',[topo,clean(r['RR']),parseInt(r['Nb Tronçons'])||0,clean(r['Accès Site MOE'])||'Nok',parseInt(r['Nb Tronçons Mesurés'])||0,parseInt(r['Nb tronçons Ok'])||0,parseInt(r['Nb tronçons Nok'])||0]);else mesData.push({id:mid++,topo,rr:clean(r['RR']),nbT:parseInt(r['Nb Tronçons'])||0,moe:clean(r['Accès Site MOE'])||'Nok',nbM:parseInt(r['Nb Tronçons Mesurés'])||0,ok:parseInt(r['Nb tronçons Ok'])||0,nok:parseInt(r['Nb tronçons Nok'])||0,updated_at:now()});}if(!USE_PG)saveJ('mesures',mesData);}
    const cddSheet=wb.Sheets['CDD'];if(cddSheet){const rows=XLSX.utils.sheet_to_json(cddSheet,{defval:null,range:1});result.sheets['CDD']=rows.length;if(USE_PG)await pool.query('DELETE FROM cdd');const cddData=[];let cid=1;for(const r of rows){const topo=clean(r['Référence topologie']);if(!topo||topo==='Référence topologie')continue;if(USE_PG)await pool.query('INSERT INTO cdd(lot,topo,dept,rr,sc,ns,mad,cmd,pfto,sin3,comment) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)',[parseLot(r['LOT']),topo,clean(r['Département']),clean(r['RR']),clean(r['Site\nCode 42C']||r['Site Code 42C']),clean(r['Nom site']),clean(r['Date MAD Orange \n']||r['Date MAD Orange']),clean(r['Numèro Commande']||r['Numéro Commande']),parseOui(r['Projet PFTO lancé par équipe o']||r['Projet PFTO lancé']),parseOui(r['Info maj sur SIN3 par Equipe p']||r['Info maj SIN3']),clean(r['Commentaire Equipe PROJET']||'')]);else cddData.push({id:cid++,lot:parseLot(r['LOT']),topo,dept:clean(r['Département']),rr:clean(r['RR']),sc:clean(r['Site\nCode 42C']||r['Site Code 42C']),ns:clean(r['Nom site']),mad:clean(r['Date MAD Orange \n']||r['Date MAD Orange']),cmd:clean(r['Numèro Commande']||r['Numéro Commande']),pfto:parseOui(r['Projet PFTO lancé par équipe o']||r['Projet PFTO lancé']),sin3:parseOui(r['Info maj sur SIN3 par Equipe p']||r['Info maj SIN3']),comment:clean(r['Commentaire Equipe PROJET']||''),updated_at:now()});}if(!USE_PG)saveJ('cdd',cddData);}
    const emplSheet=wb.Sheets['SUIVI EMPLACEMENT'];if(emplSheet){const rows=XLSX.utils.sheet_to_json(emplSheet,{defval:null,range:1});result.sheets['SUIVI EMPLACEMENT']=rows.length;if(USE_PG)await pool.query('DELETE FROM emplacements');const emplData=[];let eid=1;for(const r of rows){const topo=clean(r['Topo']);if(!topo||topo==='Topo')continue;if(USE_PG)await pool.query('INSERT INTO emplacements(topo,lot,site,dept,rr,ep,rdv,a_b,inst,cons,a_c) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)',[topo,parseLot(r['Lot']),clean(r['Site']),clean(r['Dept']),clean(r['RR']),parseOui(r['Equipe opérationelle prévenue']),parseOui(r['Envoi prise de RDV ']),parseOui(r['Annexe B signée']),parseOui(r['Installation des équipements']),parseOui(r['Consuel']),parseOui(r['Annexe C'])]);else emplData.push({id:eid++,topo,lot:parseLot(r['Lot']),site:clean(r['Site']),dept:clean(r['Dept']),rr:clean(r['RR']),ep:parseOui(r['Equipe opérationelle prévenue']),rdv:parseOui(r['Envoi prise de RDV ']),aB:parseOui(r['Annexe B signée']),inst:parseOui(r['Installation des équipements']),cons:parseOui(r['Consuel']),aC:parseOui(r['Annexe C']),updated_at:now()});}if(!USE_PG)saveJ('emplacements',emplData);}
    result.warnings=result.warnings.slice(0,10);
    res.json(result);
  } catch(e){ console.error('Import error:',e); res.status(500).json({error:e.message}); }
});

// ═══════════════════════════════════════════════
// IMPORT CPM
// ═══════════════════════════════════════════════
app.post('/api/import/cpm', requireAuth, upload.single('file'), async (req,res) => {
  if (!req.file) return res.status(400).json({error:'Aucun fichier reçu'});
  try {
    const wb=XLSX.read(req.file.buffer,{type:'buffer'});
    const cpmSheet=wb.Sheets['CPM'];if(!cpmSheet)return res.status(400).json({error:"Onglet 'CPM' introuvable"});
    const cpmRaw=XLSX.utils.sheet_to_json(cpmSheet,{header:1,defval:null});
    const synth=cpmRaw[1]||[];
    const topo=clean(synth[1]),ref_commande=clean(synth[2]),date_crmad=synth[3]?clean(String(synth[3])).slice(0,10):'';
    if(!topo)return res.status(400).json({error:'Topologie introuvable (B2)'});
    const segments=[];for(let i=5;i<cpmRaw.length;i++){const r=cpmRaw[i];if(!r[1]||r[1]==='Reference Regroupement')continue;segments.push({ref_seg:clean(r[4]),sous_ref:clean(r[2]),date_mesc:clean(r[3])?clean(String(r[3])).slice(0,10):'',type:clean(r[5]),gtr:clean(r[6]),longueur:parseInt(r[7])||0,ext_a:clean(r[10]),ext_b:clean(r[12])});}
    const segsCPM=segments.filter(s=>s.type&&s.type.startsWith('CPM_'));
    const longueur=segsCPM.reduce((a,s)=>a+s.longueur,0);
    const sites=[];const rec=wb.Sheets['RecapMAD'];if(rec){const rr=XLSX.utils.sheet_to_json(rec,{header:1,defval:null});for(let i=1;i<rr.length;i++){const code=clean(rr[i][0]),fo=parseInt(rr[i][1])||0;if(code&&!code.includes('BPU')&&fo>0)sites.push({code,nb_fo:fo});}}
    const nb_fo=sites.reduce((a,s)=>a+s.nb_fo,0);
    const hepoc=[];const hep=wb.Sheets['HEPOC-Empl_Ener'];if(hep){XLSX.utils.sheet_to_json(hep,{defval:null}).forEach(r=>{if(!r['Code_site'])return;hepoc.push({code_site:clean(r['Code_site']),nom_site:clean(r['nom_site']),cp:clean(r['lbetablissementcdpostal']),type_site:clean(r['type_site']),date_mes:clean(r['datemescommercial'])?clean(String(r['datemescommercial'])).slice(0,10):'',cmd:clean(r['nocommandefci']),gti:clean(r['GTI']),type_acces:clean(r['type_acces']),etat_baie1:clean(r['cdetatprestation_baie1'])});});}
    const lom=[];const lomS=wb.Sheets['HEPOC-LOM'];if(lomS){XLSX.utils.sheet_to_json(lomS,{defval:null}).forEach(r=>{if(!r['Code_site'])return;lom.push({code_site:clean(r['Code_site']),nom_site:clean(r['nom_site']),cp:clean(r['lbetablissementcdpostal']),date_mes:clean(r['datemescommercial'])||'',cmd:clean(r['nocommandefci']),id_lom:clean(r['idprestation_LOM']),etat:clean(r['cdetatprestation']),id_cable:clean(r['id_cable']),salle:clean(r['salle_RO_LOM_cable_ORT'])});});}
    const loir=[];const loirS=wb.Sheets['HEPOC-LOIR'];if(loirS){XLSX.utils.sheet_to_json(loirS,{defval:null}).forEach(r=>{if(!r['Code_site'])return;loir.push({code_site:clean(r['Code_site']),nom_site:clean(r['nom_site']),cp:clean(r['lbetablissementcdpostal']),date_mes:clean(r['datemescommercial'])||'',cmd:clean(r['nocommandefci']),id_loir:clean(r['idprestation_LOIR']),etat:clean(r['cdetatprestation']),ferme:clean(r['ferme_RO1']),niveau:clean(r['niveau_RO1']),connecteur:clean(r['num_connecteur_RO1'])});});}
    const pen=[];const penS=wb.Sheets['HEPOC-penetration_cable'];if(penS){XLSX.utils.sheet_to_json(penS,{defval:null}).forEach(r=>{if(!r['Code_site'])return;pen.push({code_site:clean(r['Code_site']),cp:clean(r['lbetablissementcdpostal']),cmd:clean(r['nocommandefci']),etat_ligne:clean(r['cdetatligne']),id_pen:clean(r['idprestation_PENETRATION_CABLE']),etat:clean(r['cdetatprestation'])});});}
    if(USE_PG){await pool.query(`INSERT INTO cpm(topo,ref_client,ref_commande,date_crmad,statut,nb_segments,nb_sites,nb_fo_total,longueur_totale,segments,sites,hepoc,lom,loir,penetration,date_reception) VALUES($1,$2,$3,$4,'CRMAD reçu',$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15) ON CONFLICT(topo) DO UPDATE SET ref_commande=$3,date_crmad=$4,nb_segments=$5,nb_sites=$6,nb_fo_total=$7,longueur_totale=$8,segments=$9,sites=$10,hepoc=$11,lom=$12,loir=$13,penetration=$14,imported_at=NOW(),updated_at=NOW()`,[topo,topo,ref_commande,date_crmad,segsCPM.length,sites.length,nb_fo,longueur,JSON.stringify(segments.slice(0,200)),JSON.stringify(sites),JSON.stringify(hepoc),JSON.stringify(lom),JSON.stringify(loir),JSON.stringify(pen),now().slice(0,10)]);await pool.query("UPDATE bus SET etat=GREATEST(etat,4),updated_at=NOW() WHERE id=$1 AND etat<4",[topo]);}
    else{const db=loadJ('cpm');const idx=db.findIndex(c=>c.topo===topo);const row={topo,ref_client:topo,ref_commande,date_crmad,statut:'CRMAD reçu',commentaire:'',date_reception:now().slice(0,10),anomalies:'',nb_segments:segsCPM.length,nb_sites:sites.length,nb_fo_total:nb_fo,longueur_totale:longueur,segments:segments.slice(0,200),sites,hepoc,lom,loir,penetration,imported_at:now(),updated_at:now()};if(idx>=0){row.statut=db[idx].statut||row.statut;row.commentaire=db[idx].commentaire||'';row.anomalies=db[idx].anomalies||'';db[idx]=row;}else db.push(row);saveJ('cpm',db);const buses=loadJ('bus');const bidx=buses.findIndex(b=>b.id===topo);if(bidx>=0&&buses[bidx].etat<4){buses[bidx].etat=4;saveJ('bus',buses);}}
    res.json({ok:true,topo,nb_segments:segsCPM.length,nb_sites:sites.length,nb_fo_total:nb_fo,longueur_totale:longueur});
  } catch(e){console.error('CPM import error:',e);res.status(500).json({error:e.message});}
});


// ═══════════════════════════════════════════════
// BUS : CREATE + DELETE
// ═══════════════════════════════════════════════
app.post('/api/bus', requireAuth, async (req,res) => {
  try {
    const b = req.body;
    if (!b.id?.trim()) return res.status(400).json({error:'Le nom de topologie est obligatoire'});
    if (!b.lot?.trim()) return res.status(400).json({error:'Le lot est obligatoire'});
    if (USE_PG) {
      try {
        const {rows} = await pool.query(
          `INSERT INTO bus(id,lot,dor,ext,etat,rr,av,liv,mes,risque,comment)
           VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING *`,
          [b.id.trim(),b.lot.trim(),b.dor||'',b.ext||'',b.etat??0,b.rr||'',b.av??0,b.liv||'',b.mes||'',b.risque||'Faible',b.comment||'']
        );
        return res.status(201).json(rows[0]);
      } catch(e) {
        if(e.code==='23505') return res.status(400).json({error:'Cette topologie existe déjà'});
        throw e;
      }
    }
    const db=loadJ('bus');
    if(db.find(x=>x.id===b.id.trim())) return res.status(400).json({error:'Cette topologie existe déjà'});
    const row={id:b.id.trim(),lot:b.lot.trim(),dor:b.dor||'',ext:b.ext||'',etat:b.etat??0,rr:b.rr||'',av:b.av??0,liv:b.liv||'',mes:b.mes||'',risque:b.risque||'Faible',comment:b.comment||'',created_at:now(),updated_at:now()};
    db.push(row); saveJ('bus',db); res.status(201).json(row);
  } catch(e){res.status(500).json({error:e.message});}
});

app.delete('/api/bus/:id', requireAuth, async (req,res) => {
  try {
    if(USE_PG){
      await pool.query('DELETE FROM delivery WHERE topo=$1',[req.params.id]);
      await pool.query('DELETE FROM mesures WHERE topo=$1',[req.params.id]);
      await pool.query('DELETE FROM cdd WHERE topo=$1',[req.params.id]);
      await pool.query('DELETE FROM emplacements WHERE topo=$1',[req.params.id]);
      await pool.query('DELETE FROM cpm WHERE topo=$1',[req.params.id]);
      const {rows}=await pool.query('DELETE FROM bus WHERE id=$1 RETURNING id',[req.params.id]);
      if(!rows[0]) return res.status(404).json({error:'BUS introuvable'});
      return res.json({ok:true});
    }
    ['delivery','mesures','cdd','emplacements','cpm'].forEach(t=>{const d=loadJ(t).filter(x=>x.topo!==req.params.id);saveJ(t,d);});
    const db=loadJ('bus'); const idx=db.findIndex(b=>b.id===req.params.id);
    if(idx<0) return res.status(404).json({error:'BUS introuvable'});
    db.splice(idx,1); saveJ('bus',db); res.json({ok:true});
  } catch(e){res.status(500).json({error:e.message});}
});

// ═══════════════════════════════════════════════
// DELIVERY : CREATE
// ═══════════════════════════════════════════════
app.post('/api/delivery', requireAuth, async (req,res) => {
  try {
    const f=req.body;
    if(!f.topo?.trim()) return res.status(400).json({error:'Topologie obligatoire'});
    if(USE_PG){
      await pool.query(`INSERT INTO delivery(topo,lot,etat,h_rcv,h_f,h_rnv,c_rcv,c_f,c_rnv,v_o,cf)
        VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) ON CONFLICT(topo) DO UPDATE SET
        lot=$2,etat=$3,h_rcv=$4,h_f=$5,h_rnv=$6,c_rcv=$7,c_f=$8,c_rnv=$9,v_o=$10,cf=$11,updated_at=NOW()`,
        [f.topo,f.lot||'',f.etat||'',f.hRcv||'NON',f.hF||'NON',f.hRnv||'NON',f.cRcv||'NON',f.cF||'NON',f.cRnv||'NON',f.vO||'NON',f.cf||'NON']);
      return res.status(201).json({ok:true});
    }
    const db=loadJ('delivery'); const idx=db.findIndex(d=>d.topo===f.topo);
    const row={topo:f.topo,lot:f.lot||'',etat:f.etat||'',hRcv:f.hRcv||'NON',hF:f.hF||'NON',hRnv:f.hRnv||'NON',cRcv:f.cRcv||'NON',cF:f.cF||'NON',cRnv:f.cRnv||'NON',vO:f.vO||'NON',cf:f.cf||'NON',updated_at:now()};
    if(idx>=0)db[idx]=row;else db.push(row); saveJ('delivery',db); res.status(201).json(row);
  } catch(e){res.status(500).json({error:e.message});}
});

// ═══════════════════════════════════════════════
// MESURES : CREATE + DELETE
// ═══════════════════════════════════════════════
app.post('/api/mesures', requireAuth, async (req,res) => {
  try {
    const f=req.body;
    if(!f.topo?.trim()) return res.status(400).json({error:'Topologie obligatoire'});
    if(USE_PG){
      const {rows}=await pool.query('INSERT INTO mesures(topo,rr,nb_t,moe,nb_m,ok,nok) VALUES($1,$2,$3,$4,$5,$6,$7) RETURNING *',
        [f.topo,f.rr||'',f.nbT||0,f.moe||'Nok',f.nbM||0,f.ok||0,f.nok||0]);
      return res.status(201).json({...rows[0],nbT:rows[0].nb_t,nbM:rows[0].nb_m});
    }
    const db=loadJ('mesures'); const id=Date.now();
    const row={id,topo:f.topo,rr:f.rr||'',nbT:f.nbT||0,moe:f.moe||'Nok',nbM:f.nbM||0,ok:f.ok||0,nok:f.nok||0,updated_at:now()};
    db.push(row); saveJ('mesures',db); res.status(201).json(row);
  } catch(e){res.status(500).json({error:e.message});}
});

app.delete('/api/mesures/:id', requireAuth, async (req,res) => {
  try {
    if(USE_PG){const {rows}=await pool.query('DELETE FROM mesures WHERE id=$1 RETURNING id',[req.params.id]);if(!rows[0])return res.status(404).json({error:'Mesure introuvable'});return res.json({ok:true});}
    const db=loadJ('mesures');const idx=db.findIndex(m=>m.id===+req.params.id);
    if(idx<0)return res.status(404).json({error:'Mesure introuvable'});
    db.splice(idx,1);saveJ('mesures',db);res.json({ok:true});
  } catch(e){res.status(500).json({error:e.message});}
});

// ═══════════════════════════════════════════════
// CDD : CREATE + DELETE
// ═══════════════════════════════════════════════
app.post('/api/cdd', requireAuth, async (req,res) => {
  try {
    const f=req.body;
    if(!f.topo?.trim()) return res.status(400).json({error:'Topologie obligatoire'});
    if(USE_PG){
      const {rows}=await pool.query('INSERT INTO cdd(lot,topo,dept,rr,sc,ns,mad,cmd,pfto,sin3,comment) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING *',
        [f.lot||'',f.topo,f.dept||'',f.rr||'',f.sc||'',f.ns||'',f.mad||'',f.cmd||'',f.pfto||'NON',f.sin3||'NON',f.comment||'']);
      return res.status(201).json(rows[0]);
    }
    const db=loadJ('cdd'); const id=Date.now();
    const row={id,lot:f.lot||'',topo:f.topo,dept:f.dept||'',rr:f.rr||'',sc:f.sc||'',ns:f.ns||'',mad:f.mad||'',cmd:f.cmd||'',pfto:f.pfto||'NON',sin3:f.sin3||'NON',comment:f.comment||'',updated_at:now()};
    db.push(row);saveJ('cdd',db);res.status(201).json(row);
  } catch(e){res.status(500).json({error:e.message});}
});

app.delete('/api/cdd/:id', requireAuth, async (req,res) => {
  try {
    if(USE_PG){const {rows}=await pool.query('DELETE FROM cdd WHERE id=$1 RETURNING id',[req.params.id]);if(!rows[0])return res.status(404).json({error:'CDD introuvable'});return res.json({ok:true});}
    const db=loadJ('cdd');const idx=db.findIndex(c=>c.id===+req.params.id);
    if(idx<0)return res.status(404).json({error:'CDD introuvable'});
    db.splice(idx,1);saveJ('cdd',db);res.json({ok:true});
  } catch(e){res.status(500).json({error:e.message});}
});

// ═══════════════════════════════════════════════
// EMPLACEMENTS : CREATE + PUT + DELETE
// ═══════════════════════════════════════════════
app.post('/api/emplacements', requireAuth, async (req,res) => {
  try {
    const f=req.body;
    if(!f.topo?.trim()) return res.status(400).json({error:'Topologie obligatoire'});
    if(USE_PG){
      const {rows}=await pool.query('INSERT INTO emplacements(topo,lot,site,dept,rr,ep,rdv,a_b,inst,cons,a_c) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING *',
        [f.topo,f.lot||'',f.site||'',f.dept||'',f.rr||'',f.ep||'NON',f.rdv||'NON',f.aB||'NON',f.inst||'NON',f.cons||'NON',f.aC||'NON']);
      return res.status(201).json({...rows[0],aB:rows[0].a_b,aC:rows[0].a_c});
    }
    const db=loadJ('emplacements');const id=Date.now();
    const row={id,topo:f.topo,lot:f.lot||'',site:f.site||'',dept:f.dept||'',rr:f.rr||'',ep:f.ep||'NON',rdv:f.rdv||'NON',aB:f.aB||'NON',inst:f.inst||'NON',cons:f.cons||'NON',aC:f.aC||'NON',updated_at:now()};
    db.push(row);saveJ('emplacements',db);res.status(201).json(row);
  } catch(e){res.status(500).json({error:e.message});}
});

app.delete('/api/emplacements/:id', requireAuth, async (req,res) => {
  try {
    if(USE_PG){const {rows}=await pool.query('DELETE FROM emplacements WHERE id=$1 RETURNING id',[req.params.id]);if(!rows[0])return res.status(404).json({error:'Emplacement introuvable'});return res.json({ok:true});}
    const db=loadJ('emplacements');const idx=db.findIndex(e=>e.id===+req.params.id);
    if(idx<0)return res.status(404).json({error:'Emplacement introuvable'});
    db.splice(idx,1);saveJ('emplacements',db);res.json({ok:true});
  } catch(e){res.status(500).json({error:e.message});}
});

// ═══════════════════════════════════════════════
// EXPORT EXCEL (génère un .xlsx depuis la base)
// ═══════════════════════════════════════════════
app.get('/api/export', requireAuth, async (req,res) => {
  try {
    let bus,delivery,mesures,cdd,emplacements;
    if(USE_PG){
      [bus,delivery,mesures,cdd,emplacements]=await Promise.all([
        pool.query('SELECT * FROM bus ORDER BY lot,id').then(r=>r.rows),
        pool.query('SELECT * FROM delivery ORDER BY lot,topo').then(r=>r.rows.map(r=>({...r,hRcv:r.h_rcv,hF:r.h_f,hRnv:r.h_rnv,cRcv:r.c_rcv,cF:r.c_f,cRnv:r.c_rnv,vO:r.v_o}))),
        pool.query('SELECT * FROM mesures ORDER BY topo').then(r=>r.rows.map(r=>({...r,nbT:r.nb_t,nbM:r.nb_m}))),
        pool.query('SELECT * FROM cdd ORDER BY lot,topo').then(r=>r.rows),
        pool.query('SELECT * FROM emplacements ORDER BY lot,topo').then(r=>r.rows.map(r=>({...r,aB:r.a_b,aC:r.a_c}))),
      ]);
    } else {
      bus=loadJ('bus');delivery=loadJ('delivery');mesures=loadJ('mesures');cdd=loadJ('cdd');emplacements=loadJ('emplacements');
    }

    const ETATS={0:'0 - Non démarré',1:'1 - Design',2:'2 - Delivery',3:'3 - Editions Plans',4:'4 - Travaux',5:'5 - Livré',6:'6 - MES'};
    const wb=XLSX.utils.book_new();

    // Etat BUS
    const busRows=bus.map(b=>({'Lot':b.lot,'DOR':b.dor,'Nom topologie':b.id,'Extrémité':b.ext,'Etat BUS':ETATS[b.etat]||b.etat,'RR':b.rr,'Risque':b.risque,'Commentaire':b.comment}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(busRows),'Etat BUS');

    // HME
    const hmeRows=bus.map(b=>({'ReferenceRegroupement':b.id,'LOT':b.lot,'Etat':ETATS[b.etat],'Semaine Prév. Livraison':b.liv,'Semaine Prév. MES':b.mes}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(hmeRows),'HME');

    // Suivi TRAVAUX
    const trxRows=bus.map(b=>({'LOT':b.lot,'Typologie':b.id,'RR':b.rr,'Etat BUS':ETATS[b.etat],'Avancement BUS':b.av+'%','Date livraison':b.liv,'Points bloquants / Risques':b.risque!=='Faible'?b.risque:'','Commentaires BE':b.comment,'Commentaires RR':''}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(trxRows),'Suivi TRAVAUX');

    // SUIVI DELIVERY
    const dlvRows=delivery.map(d=>({'LOT':d.lot,'Projet':d.topo,'Etat':d.etat,'HEPOC Transmis Orange':d.hRcv,'HEPOC Envoye Orange':d.hRnv,'CPM Transmis Orange':d.cRcv,'CPM Envoye Orange':d.cRnv,'Commande ferme':d.cf}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(dlvRows),'SUIVI DELIVERY');

    // Mesures
    const mesRows=mesures.map(m=>({'RR':m.rr,'Topo':m.topo,'Nb Tronçons':m.nbT,'Accès Site MOE':m.moe,'Nb Tronçons Mesurés':m.nbM,'Nb tronçons Ok':m.ok,'Nb tronçons Nok':m.nok}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(mesRows),'Mesures');

    // CDD
    const cddRows=cdd.map(c=>({'LOT':c.lot,'Reference topologie':c.topo,'Departement':c.dept,'RR':c.rr,'Site Code 42C':c.sc,'Nom site':c.ns,'Date MAD Orange':c.mad,'Numero Commande':c.cmd,'PFTO':c.pfto,'SIN3':c.sin3,'Commentaire':c.comment}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(cddRows),'CDD');

    // SUIVI EMPLACEMENT
    const emplRows=emplacements.map(e=>({'Topo':e.topo,'Lot':e.lot,'Site':e.site,'Dept':e.dept,'RR':e.rr,'Equipe opérationelle prévenue':e.ep,'Envoi prise de RDV ':e.rdv,'Annexe B signée':e.aB,'Installation des équipements':e.inst,'Consuel':e.cons,'Annexe C':e.aC}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(emplRows),'SUIVI EMPLACEMENT');

    const buf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
    const date=new Date().toISOString().slice(0,10).replace(/-/g,'');
    res.setHeader('Content-Disposition',`attachment; filename="TransProd_Export_${date}.xlsx"`);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch(e){console.error('Export error:',e);res.status(500).json({error:e.message});}
});
// ── SETUP & HEALTH ────────────────────────────────
app.get('/api/setup', async (req,res) => {
  if (!USE_PG) return res.json({status:'ok',mode:'json',message:'Mode JSON, pas de setup requis'});
  try {
    await initDB();
    await initUsers();
    const {rows}=await pool.query(`SELECT table_name FROM information_schema.tables WHERE table_schema='public' ORDER BY table_name`);
    res.json({status:'ok',tables:rows.map(r=>r.table_name),message:'Tables créées avec succès'});
  } catch(e){res.status(500).json({status:'error',error:e.message});}
});

app.get('/health', async (req,res) => {
  try {
    if(USE_PG)await pool.query('SELECT 1');
    res.json({status:'ok',mode:USE_PG?'postgresql':'json',db:USE_PG?'connected':'json-files'});
  } catch(e){res.status(500).json({status:'error',error:e.message});}
});

// ── SPA : /login et tout le reste → index.html ───
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));

// ── START ─────────────────────────────────────────
app.listen(PORT,'0.0.0.0',async()=>{
  console.log(`\n✅  TransProd → http://localhost:${PORT}`);
  console.log(`📊  Stockage   → ${USE_PG?'PostgreSQL':'JSON'}`);
  if(USE_PG){
    await initDB().catch(e=>console.error('DB init error:',e));
    await initUsers().catch(e=>console.error('Users init error:',e));
  } else {
    await initUsers().catch(e=>console.error('Users init error:',e));
  }
});
