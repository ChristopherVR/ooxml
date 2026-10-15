# Apportez votre propre serveur

OpenTeams n’a pas de service hébergé. L’espace de travail fonctionne soit **sans serveur**
(`mode: 'local'`, uniquement entre les onglets d’un même navigateur), soit avec **un serveur que
vous exploitez** et qui respecte trois petits contrats. L’implémentation de référence,
`openteams-server`, remplit les trois dans un seul processus Node d’environ 300 lignes ; tout autre
serveur qui parle les mêmes contrats fonctionne tout aussi bien.

::: danger Pas de chiffrement de bout en bout
Les discussions transitent sous forme de mises à jour Yjs par votre serveur, qui stocke en clair le
document de la salle et chaque fichier envoyé, et peut les lire. Le flux média des appels est
chiffré entre pairs (DTLS-SRTP), et un relais TURN ne peut pas le déchiffrer, mais la signalisation
passe par le serveur.
:::

## Exécuter le serveur de référence

Nécessite Node 22 ou ultérieur.

```bash
npx openteams-server                    # listens on 127.0.0.1:8787
TEAMS_TOKEN=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
```

Il se configure uniquement par l’environnement ; la commande n’accepte aucun indicateur.

| Variable        | Valeur par défaut | Signification                                                                                                                                               |
| --------------- | ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`          | `8787`            | Port d’écoute.                                                                                                                                              |
| `HOST`          | `127.0.0.1`       | Adresse d’écoute. N’utilisez `0.0.0.0` que derrière TLS.                                                                                                    |
| `TEAMS_TOKEN`   | aucune            | Si définie, les sockets la reçoivent via `?token=` et les requêtes de fichiers via `Authorization: Bearer`.                                                 |
| `TEAMS_ORIGINS` | aucune            | `Origin` autorisées pour les sockets et les fichiers, séparées par des virgules. Sans valeur : toute origine avec un jeton, loopback uniquement sans jeton. |
| `TEAMS_DATA`    | `./data`          | Répertoire de données : instantanés `rooms/<room>.bin` et `files/<workspace>/`.                                                                             |

Ou l’intégrer :

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({
	token: process.env.TEAMS_TOKEN,
	origins: ['https://app.example.com'],
	dataDir: '/var/lib/openteams',
	maxFileBytes: 25 * 1024 * 1024, // default
	log: console.log,
});
await server.listen(8787, '127.0.0.1');
// server.http is the node:http server; server.close() flushes every room to disk.
```

Faites ensuite pointer le client vers lui :

```ts
const config = {
	mode: 'server',
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	token: 'change-me', // only if TEAMS_TOKEN is set
};
```

[Déployer le serveur de référence](/fr/deploy) couvre TLS, le jeton, TURN, Docker et un exemple Caddy.

## Points d’accès

| Point d’accès                         | Ce qu’il fait                                                                                                                              |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `ws /sync/<room>`                     | Synchronisation compatible y-websocket : un `Y.Doc` et une awareness par salle, persistés dans `rooms/<room>.bin`. Messages jusqu’à 32 Mo. |
| `ws /signal/<room>`                   | Relais JSON pour la signalisation des appels : transmet chaque message texte aux autres sockets de la salle. Messages jusqu’à 256 Ko.      |
| `POST /files/<workspace>/<name>`      | Stocke une pièce jointe (plafond de 25 Mo) et répond `201 { url, size }`.                                                                  |
| `GET /files/<workspace>/<name>`       | La télécharge en pièce jointe (`Content-Disposition: attachment`, `Referrer-Policy: no-referrer`, `nosniff`).                              |
| `POST /files/link/<workspace>/<name>` | Renvoie `{ url }` : un lien signé pour ce seul fichier, valable cinq minutes, de sorte que le jeton n’apparaisse jamais dans une URL.      |
| `GET /health`                         | `{ ok, rooms, calls, auth }`.                                                                                                              |

Les noms de salle doivent correspondre à `^[\w-]{1,128}$` et les identifiants d’espace de travail
dans les chemins de fichiers à `^[\w-]{1,64}$` ; les noms de fichiers sont assainis pour ne garder
que lettres, chiffres, `.`, `-`, `_`, espaces et parenthèses, et sont limités à 120 caractères.

## Les contrats

Tout serveur qui remplit ces trois fonctions fonctionne avec toutes les intégrations.

### Synchronisation

`ws(s)://host/<base>/<room>` parle le format de fil y-websocket (étapes de synchronisation 1 et 2,
mises à jour, awareness, requête d’awareness). Un serveur y-websocket standard convient. Un serveur
qui conserve un `Y.Doc` par salle et le persiste donne un historique à ceux qui rejoignent plus
tard.

### Signalisation

`ws(s)://host/<base>/<room>` transmet chaque message texte JSON aux **autres** sockets de la salle,
et leur envoie `{ "type": "bye", "from": <id> }` lorsqu’un socket se ferme. Un socket parle pour un
seul identifiant de pair (le premier `from` qu’il utilise). Le serveur n’a pas besoin de comprendre
les messages.

### Fichiers

Facultatif. `POST|GET <origin>/files/<workspace>/<name>` avec le jeton dans un en-tête
`Authorization: Bearer`, plus `POST <origin>/files/link/<workspace>/<name>` qui renvoie un `{ url }`
signé de courte durée pour ouvrir un fichier dans un onglet du navigateur. L’origine est tirée de
`syncUrl`. Sans serveur de fichiers, donnez à l’élément une fonction `uploadFile` (pour placer les
pièces jointes dans votre propre stockage d’objets), sinon les pièces jointes ne sont partagées que
par leur nom.

Les deux sockets reçoivent `?token=<token>` lorsqu’un jeton est configuré, car un navigateur ne
peut pas définir d’en-têtes sur un WebSocket. Appliquez-le, ainsi qu’une liste d’origines
autorisées, côté serveur.

## Appels : STUN et TURN

`iceServers` transporte les URL STUN et TURN. Un STUN public suffit sur les réseaux ouverts et ne
suffit jamais derrière un NAT symétrique ou des pare-feu stricts : exploitez votre propre TURN (par
exemple coturn) et fournissez des identifiants aux clients. Les entrées TURN sans identifiants sont
rejetées par `parseServerConfig` du cœur. Des identifiants TURN statiques sont visibles par
quiconque peut ouvrir l’application ; pour tout usage public, émettez des identifiants de courte
durée. Le [guide de déploiement](/fr/deploy#appels-a-travers-les-nat-turn) contient un exemple
coturn.

## Ce que le serveur de référence ne fait pas

Soyez au clair sur ces points avant d’y faire venir des personnes :

- **Pas de chiffrement de bout en bout.** Il stocke et peut lire chaque message et chaque fichier.
- **Pas de comptes utilisateur.** `TEAMS_TOKEN` est un secret unique partagé par tous, conservé dans
  le navigateur. Il tient les inconnus à l’écart du serveur d’une petite équipe ; ce n’est pas une
  authentification par utilisateur. Les contrôles d’auteur et de rôle dans le client peuvent être
  contournés par un client hostile, donc le serveur est la frontière de confiance : placez votre
  propre authentification devant (un proxy qui vérifie une session) pour un usage réel.
- **Pas de TLS propre.** Exécutez-le derrière un proxy inverse qui termine TLS et met à niveau les
  WebSockets ; les navigateurs refusent `ws://` depuis une page `https://`.
- **Un seul processus, salles en mémoire.** Il ne passe pas à l’échelle horizontalement. Une salle
  est déchargée 30 secondes après la fermeture de son dernier socket.
- **Un instantané par salle, sans compactage.** Le document grossit avec l’historique et n’est
  jamais élagué.
- **Stockage de fichiers brut.** Pas de liste, pas de suppression, pas de quota par espace de
  travail ; envoyer un fichier dont le nom existe déjà le remplace. Les fichiers sont limités à
  25 Mo chacun.
- **Pas de limitation de débit** et pas de serveur média : les appels sont de pair à pair
  (maillage).
