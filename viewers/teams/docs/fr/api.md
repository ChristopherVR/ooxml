# API de l’élément et événements

Toutes les intégrations enveloppent le même élément personnalisé `<teams-app>`. Elles transmettent ces propriétés et réémettent ces événements dans l’idiome de leur framework (voir les [guides par framework](/frameworks/react)).

## Attributs et propriétés

| Propriété       | Attribut        | Description                                                                                                                    |
| --------------- | --------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `workspaceId`   | `workspace-id`  | La salle partagée : 1 à 128 caractères parmi lettres, chiffres, `-` et `_`.                                                    |
| `userName`      | `user-name`     | Nom affiché.                                                                                                                   |
| `userId`        | `user-id`       | Identifiant stable ; par défaut, un identifiant mémorisé dans ce navigateur.                                                   |
| `config`        | `server-config` | `{ mode: 'local' \| 'server', syncUrl?, signalingUrl?, iceServers, iceTransportPolicy?, token? }`. L’attribut accepte du JSON. |
| `uploadFile`    |                 | La destination des pièces jointes ; par défaut, le point d’accès `/files` du serveur.                                          |
| `openers`       |                 | Gestionnaires par type Office (`docx`, `xlsx`, `pptx`, `vsdx`).                                                                |
| `embeds`        |                 | Rappels par type de fichier renvoyant l’URL d’une page de visionneuse compatible avec le bac à sable ; utile pour PowerPoint.  |
| `client` (read) |                 | Le client du cœur derrière l’élément (état en entrée, actions en sortie).                                                      |

Si vous omettez `config`, l’élément propose une boîte de dialogue de paramètres (mémorisée dans ce navigateur) ou s’exécute en mode local.

## Aperçus de contenu

L’ouverture d’une pièce jointe affiche par défaut un aperçu dans l’espace de travail. Word et Excel
utilisent les éléments d’édition natifs, d’abord en lecture seule ; Visio utilise sa visionneuse.
Markdown dispose d’un moteur de rendu à sous-ensemble sûr ; les fichiers texte affichent le texte
littéral ; HTML et les sites web utilisent des cadres en bac à sable. `openers` et l’événement
annulable `teams-open-file` restent prioritaires.

PowerPoint exige une page d’intégration fournie par l’hôte :

```ts
const embeds = {
	pptx: ({ url }: { url: string | undefined }) =>
		`https://viewer.example.com/embed?source=${encodeURIComponent(url ?? '')}`,
};
```

Transmettez `embeds` à n’importe quelle intégration de composant. La page doit comprendre l’URL du
fichier et prendre en charge un bac à sable avec scripts et formulaires mais sans permission de même
origine. Les récupérations de fichiers omettent les identifiants ; les serveurs de fichiers
d’autres origines doivent autoriser CORS. Sans visionneuse PowerPoint configurée, le volet signale
cette limite et propose l’ouverture externe.

La méthode `previewContent({ attachment, url })` de l’élément ouvre localement un contenu fourni
par l’hôte. Pour un site web, utilisez `attachment: { name: 'Project site', kind:
'other', mime: 'text/html' }`. Cela ne crée pas d’onglet de canal partagé. La vue Fichiers propose
aussi **Aperçu du site web** (**Preview website**). Les sites qui refusent l’intégration en cadre
peuvent être ouverts à l’extérieur.

## Onglets de canal partagés et copies Excel

Utilisez **Pin as tab** à côté d’un fichier envoyé, ou **Add tab** pour ajouter un site web nommé.
Les définitions sont partagées ; chaque utilisateur sélectionne son propre onglet. Le créateur peut
le renommer ou le supprimer grâce à des contrôles de propriété côté client. Supprimer un onglet
ailleurs laisse disponible un volet local déjà ouvert jusqu’à sa fermeture.

Excel propose l’édition locale et **Save copy to channel** lorsque le stockage de fichiers est
configuré. Chaque enregistrement envoie une copie au nom unique et la publie dans le canal d’origine,
en préservant l’original. Les enregistrements en échec conservent les modifications ; la navigation
demande confirmation avant d’abandonner un travail non enregistré et attend la fin d’un
enregistrement en cours. Ce n’est ni de la co-édition ni une réécriture versionnée. Les hôtes qui
remplacent de force l’espace de travail ou l’élément doivent gérer eux-mêmes la confirmation du
travail non enregistré.

Actions brutes du client :

```ts
const tab = teams.addTab('Budget', { type: 'file', attachment });
teams.addTab('Project site', { type: 'website', url: 'https://example.com' });
teams.renameTab(tab!.id, 'Budget review');
teams.removeTab(tab!.id);
await teams.saveFileCopy(channelId, file);
```

`TeamsState.tabs` contient les définitions du canal sélectionné ; `canUploadFiles` indique qu’un
stockage est configuré. `OpenFileDetail.channelId` enregistre le canal capturé à l’ouverture.
`previewContent` accepte aussi ce champ lorsqu’un hôte souhaite activer les copies enregistrées
dans le canal. Les contrôles de propriété exigent une autorisation côté serveur avant de pouvoir
être traités comme un contrôle d’accès.

## Création, envoi et recherche de fichiers

Les deux vues Fichiers proposent l’envoi, la création d’un classeur Excel vierge et la recherche.
Les envois et les nouveaux classeurs ciblent le canal sélectionné, nommé à côté des commandes. Ils
exigent un stockage configuré ; les échecs peuvent être réessayés dans le canal d’origine. Les
fichiers de même nom utilisent des URL de stockage distinctes. Les envois acceptent jusqu’à 20
fichiers de 32 Mio chacun (le plafond propre au serveur de référence est de 25 Mo par fichier, sauf si vous relevez `maxFileBytes`) ; l’interface signale l’activité mais n’offre ni progression en octets ni
annulation.

```ts
await teams.uploadFiles(channelId, files);
const workbook = await teams.createWorkbook(channelId, 'Budget'); // Budget.xlsx
```

La création publie une pièce jointe adossée à un vrai fichier XLSX. La recherche filtre les
métadonnées chargées par nom de fichier, auteur, canal et type reconnu. Elle ne cherche pas à
l’intérieur du contenu des fichiers. Ces actions capturent explicitement le canal avant la création
ou l’envoi asynchrone, de sorte que la sélection d’un autre canal ne peut pas rediriger le
résultat.

## Événements

| Événement             | Signification                                                                                         |
| --------------------- | ----------------------------------------------------------------------------------------------------- |
| `teams-ready`         | L’élément est prêt.                                                                                   |
| `teams-open-file`     | Un utilisateur a ouvert un fichier joint. Annulable : `preventDefault()` pour l’ouvrir vous-même.     |
| `teams-config-change` | L’utilisateur a enregistré de nouveaux paramètres de serveur dans la boîte de dialogue de paramètres. |

Notations selon l’intégration : `onReady`, `onOpenFile`, `onConfigChange` (React, Solid, Svelte, JS natif) ; `@ready`, `@open-file`, `@config-change` (Vue) ; `(ready)`, `(openFile)`, `(configChange)` (Angular).

## Le client brut

`createTeams(options)` (le hook brut de chaque intégration) renvoie le client du cœur :

```ts
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
teams.on('notice', (notice) => show(notice)); // errors a user should see
```

`getState()` renvoie un `TeamsState` immuable : `status` de la connexion, `channels` avec compteurs de non-lus, `messages` et `files` du canal sélectionné, `people` (présence), `typing`, l’état de la zone de saisie, les résultats de recherche et `call`. Les actions comprennent `select`, `createChannel`, `send`, `startReply`, `startEdit`, `deleteMessage`, `toggleReaction`, `notifyTyping`, `setAvailability`, `search`, `openCall`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`, `toggleScreenShare`, `toggleHand` et `destroy`. Voir [Architecture](/architecture) pour l’articulation des couches.
