# NetCollect — Guide de déploiement

Application de suivi opérationnel réseau BUS / Orange.

---

## Option 1 — Lancer en local (test rapide)

### Prérequis
- Node.js 18+ installé → https://nodejs.org

### Étapes
```bash
# 1. Décompresser le projet
unzip netcollect.zip
cd netcollect

# 2. Installer les dépendances
npm install

# 3. Lancer
npm start
```

L'application est accessible sur **http://localhost:3000**

---

## Option 2 — Déploiement équipe avec Docker (recommandé)

### Prérequis
- Docker Desktop installé → https://docker.com

### Étapes
```bash
cd netcollect
docker-compose up -d
```

L'application est accessible sur **http://VOTRE_IP:3000**

Les données sont sauvegardées dans `./data/netcollect.db`

### Arrêt / Redémarrage
```bash
docker-compose stop     # Arrêter
docker-compose start    # Redémarrer
docker-compose down     # Supprimer le conteneur (données conservées)
```

---

## Option 3 — Serveur VPS (OVH, Infomaniak, etc.)

### Sur le serveur
```bash
# Installer Node.js 20
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs

# Copier le projet
scp -r netcollect/ user@VOTRE_SERVEUR:/opt/netcollect

# Sur le serveur
cd /opt/netcollect
npm install --production
node server.js
```

### Avec PM2 (relance automatique au reboot)
```bash
npm install -g pm2
pm2 start server.js --name netcollect
pm2 startup
pm2 save
```

### Accès via domaine (nginx reverse proxy)
```nginx
server {
    listen 80;
    server_name netcollect.votre-domaine.fr;

    location / {
        proxy_pass http://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
    }
}
```

---

## Option 4 — Railway.app (hébergement cloud gratuit)

1. Créer un compte sur https://railway.app
2. Nouveau projet → "Deploy from GitHub"
3. Uploader le code ou connecter GitHub
4. Railway détecte Node.js automatiquement
5. Ajouter un volume persistant pour `/app/data`
6. URL publique générée automatiquement

---

## Utilisation quotidienne

### Import Excel
1. Aller sur la page **Import Excel**
2. Glisser-déposer `SUIVI_PROJET_Collecte.xlsx`
3. Vérifier le rapport d'import
4. Cliquer "Voir les BUS importés"

> ✅ Les commentaires et modifications manuelles sont préservés lors des réimports.

### Modifier un BUS
- Cliquer sur n'importe quelle ligne du tableau BUS
- Onglet **Synthèse** : changer l'état, l'avancement, le commentaire
- Onglet **Delivery** : mettre à jour HEPOC / CPM
- Onglet **Historique** : voir toutes les modifications

### Identifiant utilisateur
- Cliquer "Changer de nom" en bas de la sidebar
- Saisir son prénom ou identifiant
- Toutes les modifications sont tracées avec ce nom

---

## Structure du projet

```
netcollect/
├── server.js           ← Backend Express + routes API
├── package.json
├── Dockerfile
├── docker-compose.yml
├── public/
│   └── index.html      ← Frontend (SPA vanilla JS)
└── data/               ← Créé automatiquement
    └── netcollect.db   ← Base SQLite (à sauvegarder !)
```

## API disponibles

| Méthode | URL | Description |
|---------|-----|-------------|
| GET | /api/bus | Liste tous les BUS |
| PUT | /api/bus/:id | Modifier un BUS |
| GET | /api/delivery | Données delivery |
| PUT | /api/delivery/:topo | Modifier delivery |
| GET | /api/mesures | Mesures |
| GET | /api/cdd | CDD |
| GET | /api/emplacements | Emplacements |
| GET | /api/stats | Statistiques dashboard |
| GET | /api/history/:id | Historique d'un BUS |
| GET | /api/comments/:topo | Commentaires |
| POST | /api/comments/:topo | Ajouter commentaire |
| POST | /api/import | Import fichier Excel |

## Sauvegarde

Le fichier `data/netcollect.db` contient **toutes les données**.
Faire une copie régulière de ce fichier suffit comme sauvegarde.

```bash
# Exemple sauvegarde journalière (crontab)
0 2 * * * cp /opt/netcollect/data/netcollect.db /backups/netcollect_$(date +%Y%m%d).db
```
