# Exploiter votre propre serveur

## Le serveur de référence

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable        | Signification                                                                                               |
| --------------- | ----------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`  | Adresse d’écoute. N’utilisez `0.0.0.0` que derrière TLS.                                                    |
| `TEAMS_TOKEN`   | Si définie, les sockets la reçoivent via `?token=` et les requêtes de fichiers via `Authorization: Bearer`. |
| `TEAMS_ORIGINS` | `Origin` autorisées pour les sockets et les fichiers, séparées par des virgules. Vide autorise toutes.      |
| `TEAMS_DATA`    | Répertoire de données (`./data` par défaut) : instantanés `rooms/<room>.bin` et `files/`.                   |

Il sert `ws /sync/<room>` (compatible y-websocket, un `Y.Doc` persisté par salle),
`ws /signal/<room>` (un relais JSON) et `POST|GET /files/<workspace>/<name>` (plafond de 25 Mo,
noms assainis, téléchargements envoyés en pièces jointes avec `Referrer-Policy: no-referrer`). Avec
un jeton, un navigateur ouvre les fichiers via `POST /files/link/<workspace>/<name>` (jeton dans
l’en-tête), qui renvoie un lien signé pour ce seul fichier et valable cinq minutes. `GET /health`
indique le nombre de salles et d’appels.

**Placez-le derrière TLS.** Les navigateurs refusent `ws://` depuis une page `https://`, et un jeton
ou la signalisation des appels en HTTP simple sont lisibles sur le réseau. Tout proxy inverse qui
met à niveau les WebSockets convient (nginx, Caddy, Traefik). Exemple Caddy :

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

Faites ensuite pointer l’application vers `wss://teams.example.com/sync` et
`wss://teams.example.com/signal`.

**Le jeton est un secret partagé** conservé dans le navigateur : il tient les inconnus à l’écart du
serveur d’une petite équipe, ce n’est pas une authentification par utilisateur. Pour de vrais
comptes, placez votre propre authentification devant (un proxy qui vérifie une session, ou
remplacez `tokenOk` par une vérification de jeton) et conservez les trois mêmes contrats.

## Appels à travers les NAT : TURN

Les appels sont de pair à pair. Un STUN public fonctionne sur les réseaux ouverts ; derrière un NAT
symétrique ou des pare-feu stricts, il faut un relais TURN, qui ne transporte lui aussi le flux
média que sous forme chiffrée (DTLS-SRTP). Un coturn minimal (`turnserver.conf`) :

```
listening-port=3478
tls-listening-port=5349
realm=turn.example.com
fingerprint
lt-cred-mech
user=teams:CHANGE_ME
no-multicast-peers
cert=/etc/letsencrypt/live/turn.example.com/fullchain.pem
pkey=/etc/letsencrypt/live/turn.example.com/privkey.pem
```

et, dans les paramètres du serveur de l’application, un serveur ICE par ligne :

```
stun:turn.example.com:3478
turn:turn.example.com:3478,turns:turn.example.com:5349 teams CHANGE_ME
```

Des identifiants TURN statiques sont visibles par quiconque peut ouvrir l’application. Pour tout
usage public, émettez des identifiants de courte durée (coturn `use-auth-secret` plus un petit
point d’accès qui les signe) et transmettez-les dans les `iceServers` que vous donnez à l’élément.

## Docker

```dockerfile
FROM node:22-alpine
WORKDIR /app
COPY server/package.json ./
RUN npm install --omit=dev
COPY server/index.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data
VOLUME /data
EXPOSE 8787
CMD ["node", "index.mjs"]
```

Lancez-le avec `TEAMS_TOKEN` et `TEAMS_ORIGINS` définis, derrière votre proxy TLS.

## Utiliser un autre serveur

Tout ce qui respecte les contrats décrits dans `docs/teams-area.md` du cœur convient : un serveur
`y-websocket` standard pour la synchronisation, n’importe quelle diffusion WebSocket pour la
signalisation, n’importe quel stockage d’objets pour les fichiers (donnez à `<teams-app>` une
fonction `uploadFile`, ou définissez `uploadFile` sur l’élément, et les pièces jointes y seront
envoyées à la place).
