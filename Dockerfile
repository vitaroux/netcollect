FROM node:20-alpine

WORKDIR /app

# Dépendances système pour better-sqlite3
RUN apk add --no-cache python3 make g++

COPY package*.json ./
RUN npm install --production

COPY . .

# Dossier de données (volume persistant)
RUN mkdir -p data

EXPOSE 3000

CMD ["node", "server.js"]
