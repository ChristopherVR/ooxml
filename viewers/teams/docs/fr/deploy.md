# Exploiter votre propre serveur

## Le serveur de référence

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable                                                | Signification                                                                                                                                                                                                                                                                                 |
| ------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`                                          | Adresse d’écoute. N’utilisez `0.0.0.0` que derrière TLS.                                                                                                                                                                                                                                      |
| `TEAMS_TOKEN`                                           | Si définie, les sockets la reçoivent via `?token=` et les requêtes de fichiers via `Authorization: Bearer`.                                                                                                                                                                                   |
| `TEAMS_ORIGINS`                                         | `Origin` autorisées pour les sockets et les fichiers, séparées par des virgules. Vide : toute origine si un jeton est défini ; sans jeton, seulement les origines loopback et les clients sans `Origin`.                                                                                      |
| `TEAMS_DATA`                                            | Répertoire de données (`./data` par défaut) : instantanés `rooms/<room>.bin` et `files/`.                                                                                                                                                                                                     |
| `TEAMS_TRUST_PROXY`, `--trust-proxy[=N]`                | Regroupe le limiteur de débit d’authentification et le plafond de connexions par adresse selon `X-Forwarded-For` (l’entrée du N-ième proxy de confiance, 1 par défaut). Désactivé par défaut ; ne l’activez que derrière un proxy qui définit cet en-tête, sinon un client peut le falsifier. |
| `--insecure-dev`                                        | Démarre sans comptes et sans fonction `authenticate` (option `insecureDev`). Sans cette option, le serveur ne démarre que si `TEAMS_USERS_FILE` est défini. Développement uniquement.                                                                                                         |
| `TEAMS_INSECURE`                                        | `1` (ou `--insecure`) autorise un `HOST` non loopback sans jeton. Sinon le serveur refuse de démarrer.                                                                                                                                                                                        |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Sockets ouverts au total (500 par défaut) et par adresse distante (20).                                                                                                                                                                                                                       |
| `TEAMS_MAX_ROOMS`                                       | Salles de synchronisation et d’appel (1000 chacune par défaut).                                                                                                                                                                                                                               |
| `TEAMS_MAX_DOC_MB`                                      | Plus grand document de salle persisté, en Mo (64 par défaut) ; les mises à jour au-delà sont rejetées.                                                                                                                                                                                        |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Messages par seconde (100 par défaut) et rafale (200) par socket ; un socket qui dépasse est fermé.                                                                                                                                                                                           |
| `TEAMS_FILES_QUOTA_MB`                                  | Taille totale du répertoire des fichiers, en Mo (1024 par défaut) ; les envois au-delà reçoivent 507.                                                                                                                                                                                         |

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
comptes, donnez au serveur `users` (ou `TEAMS_USERS_FILE`) ou une fonction `authenticate` pour votre fournisseur d’identité, et définissez `TEAMS_SESSION_SECRET` pour que les sessions survivent aux redémarrages (voir "Accounts and sessions" dans le README du serveur). Les trois contrats restent les mêmes.

## Modèle de sécurité

- **Un seul jeton partagé ouvre toutes les salles et tous les fichiers.** Il n’y a pas de droits par
  utilisateur ni par salle. Traitez le jeton comme le mot de passe d’une équipe.
- **Tous les membres d’une salle se font mutuellement confiance.** Le serveur ne fait que transporter
  des octets Yjs : n’importe quel membre peut modifier ou supprimer les messages des autres et falsifier
  les identifiants d’auteur et les rôles. N’invitez pas des personnes à qui vous ne confieriez pas toute
  la salle.
- **Le jeton et la session circulent dans la chaîne de requête du WebSocket** (`?token=`, `?session=`), que les proxys et les
  journaux d’accès peuvent enregistrer et que les navigateurs peuvent garder dans l’historique. Le serveur les accepte aussi
  sous forme d’entrées `Sec-WebSocket-Protocol` (`openteams.session.<session>`, `openteams.token.<token>`, à côté d’un
  `openteams` simple), qui n’apparaissent ni dans les URL ni dans les journaux ; le client fourni ne les envoie pas encore
  (travail à suivre dans la zone `teams` du cœur). Placez le serveur derrière TLS et ne journalisez pas les chaînes de
  requête. Les fichiers n’acceptent jamais le jeton dans une URL : ils utilisent l’en-tête `Authorization` ou un lien signé.
- **Sûr par défaut.** Sans comptes, sans fonction `authenticate` ni `authorizeSync`, le serveur refuse de démarrer ;
  `--insecure-dev` est l’exception explicite, assortie d’un avertissement. Un jeton partagé seul n’est pas une authentification.
- **Les valeurs par défaut sont locales.** Sans jeton, le serveur refuse un `HOST` non loopback (sauf
  avec `TEAMS_INSECURE=1`) et n’accepte que des origines loopback, si bien qu’un site web visité ne peut
  pas piloter un serveur de votre propre machine. Définissez `TEAMS_TOKEN` et `TEAMS_ORIGINS` pour tout
  déploiement partagé.
- **Utilisez des identifiants TURN de courte durée** (voir plus bas) ; les statiques sont visibles de quiconque ouvre l’application.
- **Les limites** de connexions, de salles, de taille de document, de débit de messages et de stockage
  de fichiers (les variables ci-dessus) bornent les abus, mais ne remplacent pas la limitation de débit
  d’un proxy inverse.
- Un socket ne peut parler que pour les identifiants de présence et l’identifiant de pair d’appel qu’il
  a lui-même annoncés, et un signal d’appel adressé n’est remis qu’à ce pair. Un envoi n’écrase jamais
  un fichier : un nom déjà pris est enregistré sous un nouveau, renvoyé dans le champ `url` de la réponse.

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
COPY server/bin.mjs server/index.mjs server/auth.mjs server/passwords.mjs server/guards.mjs \
  server/rate-limit.mjs server/chat-access.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data TEAMS_USERS_FILE=/config/users.json
RUN mkdir /data && chown node:node /data
VOLUME /data
EXPOSE 8787
USER node
CMD ["node", "bin.mjs"]
```

Construisez-le depuis `viewers/teams`. Il copie chaque module que le point d’entrée importe (les `files` de `server/package.json`) ; ne copier que `index.mjs` échoue avec un module manquant. La recette a été vérifiée en exécutant ces fichiers avec `npm install --omit=dev` sous Node ; l’image elle-même n’a pas été construite.

Lancez-le avec un fichier de comptes monté, `TEAMS_SESSION_SECRET` et `TEAMS_ORIGINS` définis, derrière votre proxy TLS :

```bash
docker run -d --name openteams -p 127.0.0.1:8787:8787 \
  -v openteams-data:/data -v "$PWD/users.json:/config/users.json:ro" \
  -e TEAMS_SESSION_SECRET=change-me -e TEAMS_ORIGINS=https://app.example.com \
  openteams-server
```

Sans fichier de comptes monté, le conteneur s’arrête avec le message « refusing to start » ; n’ajoutez pas `--insecure-dev` à un
conteneur joignable depuis un réseau. Ajoutez `--trust-proxy` à la commande (`CMD ["node", "bin.mjs", "--trust-proxy"]`) seulement si
un proxy inverse définit `X-Forwarded-For`. `docker stop` envoie `SIGTERM`, auquel le serveur répond en enregistrant d’abord chaque
salle (Docker attend dix secondes par défaut, ce qui suffit). L’image crée `/data` appartenant à `node` ; avec un montage de liaison,
changez le propriétaire du répertoire hôte en uid 1000 avec `chown`.

## Exploitation

### Sauvegarde et restauration

Tout ce que le serveur conserve se trouve sous `TEAMS_DATA` (`./data` par défaut) :

| Chemin                     | Contenu                                                                                                                     |
| -------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `rooms/<room>.bin`         | Un instantané Yjs par salle : tout le chat, les canaux, les calendriers et les données des onglets.                         |
| `files/<workspace>/`       | Les fichiers téléversés.                                                                                                    |
| `session-revocations.json` | Les générations de révocation des sessions (sans lui, les sessions révoquées redeviennent valides jusqu’à leur expiration). |

Conservez aussi ce qui ne se trouve pas dans le répertoire de données : votre fichier de comptes (`TEAMS_USERS_FILE`) et `TEAMS_SESSION_SECRET`.
Modifier ou perdre le secret déconnecte tout le monde, ce qui est sans gravité.

Pour sauvegarder, copiez tout le répertoire de données. Les instantanés de salles sont écrits dans un fichier temporaire puis renommés, si bien
qu’une copie ne voit jamais un instantané à moitié écrit ; en revanche, une copie faite pendant que le serveur tourne peut manquer les modifications
de la dernière demi-seconde (les écritures sont différées de 500 ms) et capturer un fichier plus récent que ceux de sa salle. Pour une copie
cohérente, arrêtez d’abord le serveur (`SIGTERM` enregistre chaque salle avant de quitter), copiez, puis redémarrez-le ; ou prenez un instantané
du volume au niveau du système de fichiers.

Pour restaurer, arrêtez le serveur, remplacez le répertoire de données par la sauvegarde (en gardant `rooms/`, `files/` et `session-revocations.json`
ensemble), restaurez le fichier de comptes et le secret, puis démarrez-le. Les clients connectés se reconnectent et fusionnent : Yjs conserve ce qu’ils
détiennent et que la sauvegarde ne contient pas, de sorte qu’une restauration n’efface pas le travail encore ouvert dans un navigateur.

### Mise à niveau

1. Lisez le `CHANGELOG.md` du serveur pour la plage de versions franchie ; les ruptures de compatibilité sont signalées.
2. Sauvegardez le répertoire de données.
3. Arrêtez le serveur (`SIGTERM`), installez la nouvelle version (`npx openteams-server@<version>`, ou reconstruisez l’image) et démarrez-le. Les instantanés de salles sont de simples mises à jour Yjs ; les plus anciens sont lus tels quels.
4. Vérifiez `GET /health`, puis ouvrez une salle.

Mise à niveau depuis une version antérieure au réglage sûr par défaut : un serveur qui démarrait sans comptes refuse désormais de le faire.
Ajoutez des comptes (`TEAMS_USERS_FILE`), un hook `authenticate` ou `authorizeSync`, ou passez `--insecure-dev` si (et seulement si) il s’agit d’un serveur de
développement local. Un `TEAMS_TOKEN` partagé seul ne démarre plus le serveur.

### Liste de contrôle pour la production

- [ ] Comptes configurés (`TEAMS_USERS_FILE` avec des identifiants hachés produits par `openteams-hash-password`, ou un hook `authenticate`) ; `--insecure-dev` n’est pas utilisé.
- [ ] `TEAMS_SESSION_SECRET` défini à une valeur aléatoire longue, conservée hors du dépôt, pour que les sessions survivent aux redémarrages.
- [ ] TLS se termine devant le serveur ; l’application utilise `wss://`. Le serveur lui-même écoute sur loopback ou sur un réseau privé.
- [ ] `TEAMS_ORIGINS` liste les origines exactes de votre application.
- [ ] `--trust-proxy` défini seulement si le proxy définit `X-Forwarded-For` et que les clients ne peuvent joindre le serveur qu’à travers lui ; sinon, laissez-le désactivé (tous les clients derrière le proxy partagent alors une même clé de limitation de débit, ce qui est sûr mais grossier).
- [ ] Les journaux d’accès du proxy n’enregistrent pas les chaînes de requête (les sockets portent `?token=` et `?session=`), ou vous l’acceptez.
- [ ] `TEAMS_ADMIN_TOKEN` défini si vous avez besoin de `POST /auth/revoke`, et gardé secret.
- [ ] Le répertoire de données est sur un volume persistant, accessible en écriture seulement à l’utilisateur du serveur, et sauvegardé.
- [ ] Limites revues (`TEAMS_MAX_*`, `TEAMS_FILES_QUOTA_MB`) et le proxy applique sa propre limitation de débit.
- [ ] Identifiants TURN de courte durée, pas de secret statique dans la page.
- [ ] Le processus est supervisé et arrêté par `SIGTERM` afin que les salles soient enregistrées.
- [ ] Vous avez lu [Limitations](./limitations.md) : pas de chiffrement de bout en bout, les membres d’une salle se font mutuellement confiance.

## Utiliser un autre serveur

Tout ce qui respecte les contrats décrits dans `docs/teams-area.md` du cœur convient : un serveur
`y-websocket` standard pour la synchronisation, n’importe quelle diffusion WebSocket pour la
signalisation, n’importe quel stockage d’objets pour les fichiers (donnez à `<teams-app>` une
fonction `uploadFile`, ou définissez `uploadFile` sur l’élément, et les pièces jointes y seront
envoyées à la place).
