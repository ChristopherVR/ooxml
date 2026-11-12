# Intégration OneDrive

Pour une configuration pas à pas (enregistrement dans Microsoft Entra, connexion et exemple complet), commencez par le [démarrage rapide OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

Pour un autre service ou une implémentation OneDrive personnalisée, voir [les intégrations de stockage indépendantes du fournisseur](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md).
`createStorageSession` se connecte à la même propriété `storage` dans tous les paquets Visio, avec les contrôles de révision gérés par l'hôte, les gestionnaires de fusion, l'interface de conflit et les politiques de partage.

## Fichier > Ouvrir avec les emplacements cloud {#file-open-with-cloud-places}

Transmettez `fileProviders` (une propriété de l'élément et une prop de chaque liaison) pour ajouter des emplacements cloud à Fichier > Ouvrir. Lorsque la liste n'est pas vide, Ouvrir affiche **Cet appareil** (la vignette habituelle Parcourir) et un emplacement par fournisseur, qui affiche le navigateur partagé `<office-ui-cloud-files>` pour les dessins `.vsdx` : connexion, Récents, Mes fichiers, recherche et ouverture depuis un lien de partage. Les libellés proviennent de la langue d'affichage de Visio dans chaque langue.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDrawingSnapshots } from 'ooxml-core/visio/collab';

const oneDrive = createOneDriveFileProvider({
	getAccessToken, // the host's own sign-in
	signIn, // opens the host's sign-in UI and resolves with the account
	canWrite: async (file) => file.driveId === (await myDriveId()),
	mergers: {
		vsdx: ({ base, local, remote }) => mergeDrawingSnapshots(base.bytes, local, remote.bytes),
	},
});
// <VisioViewer fileProviders={[oneDrive]} /> (React); [fileProviders] in Angular; :file-providers in Vue.
```

Un dessin choisi à cet endroit se charge par le chemin `.vsdx` normal et prend le nom du fichier, puis :

- Il est en lecture seule lorsque le fichier l'est (`canEdit` à false). Un hôte qui a déjà rendu le visualiseur en lecture seule (`canEdit`, `readOnly`, un rôle `viewer` de collaboration ou `capabilities.edit`) reste en lecture seule pour un fichier modifiable : c'est la plus stricte des deux restrictions qui s'applique.
- Fichier > Enregistrer et Enregistrer sous utilisent le stockage propre à ce fichier au lieu du `storage` de l'hôte, jusqu'à l'ouverture d'un autre dessin (Ouvrir, Nouveau, `load` ou `document`). Enregistrer ne se résout qu'une fois que le fournisseur a stocké les octets : le dessin apparaît alors comme enregistré, et il reste non enregistré avec un événement `document-error` si l'envoi échoue.
- L'événement annulable `file-command` se déclenche toujours en premier (`open`, `save`) ; annuler `open` laisse le dessin courant en place. Le volet des coulisses se ferme une fois le dessin ouvert.
- Les diagnostics d'export sont signalés comme pour tout autre enregistrement. Le `.vsd` ancien n'est pas une cible d'édition cloud.

Les démos utilisent par défaut un emplacement OneDrive, connecté via la boîte de dialogue réservée à la démo (un identifiant client, ou un jeton de Graph Explorer collé dans **Jeton d'accès**). Cette boîte de dialogue fait partie de la configuration de la démo, pas de l'interface du produit ; voir `demos/onedrive`. Pour passer cette boîte de dialogue, définissez `VITE_ONEDRIVE_CLIENT_ID` (et, au besoin, `VITE_ONEDRIVE_AUTHORITY`) dans `demos/visio/.env.local`, lancez `bun run dev` dans `viewers/visio`, et enregistrez l'adresse affichée (avec une barre oblique finale) comme URI de redirection d'application monopage.

## Compte > Services connectés {#account-connected-services}

Les mêmes `fileProviders` remplissent aussi la page **Compte** de la vue Fichier avec une liste « Services connectés » (le composant partagé `<office-ui-connected-services>`) : une ligne par fournisseur avec le compte connecté et les boutons Se connecter ou Se déconnecter, qui appellent les `signIn` et `signOut` du fournisseur. Un fournisseur sans `signOut` indique que la session est gérée par l'hôte. La liste suit les `subscribe` des fournisseurs : se connecter depuis Fichier > Ouvrir la met à jour, et ses libellés proviennent des dictionnaires Visio dans chaque langue. La page conserve aussi le profil local et les informations sur le produit. Comme dans Word et Excel, la page Compte n'apparaît que s'il existe des `fileProviders` ; sans eux, le profil reste modifiable dans **Fichier > Options**.

## Utiliser directement une session {#using-a-session-directly}

Visio utilise la connexion partagée `ooxml-core/embedding/onedrive` pour les fichiers VSDX natifs. L'hôte gère la connexion Microsoft, la résolution des liens et la vérification des autorisations propres à chaque élément. Transmettez les octets téléchargés au visualiseur existant, avec une restriction de lecture seule de l'hôte lorsque l'accès est inconnu :

```ts
// Set these through the element or binding's existing props/options.
const canEdit = opened.file.access === 'write';
const options = { storage: session.storage, canEdit, readOnly: !canEdit };
// After mounting with options:
await viewer.load(opened.bytes);
```

`exportVsdx()` renvoie un objet contenant les `bytes` natifs et les diagnostics du writer. Un bouton Enregistrer personnalisé peut appeler `await session.save((await viewer.exportVsdx()).bytes)`. L'enregistrement natif de Fichier utilise le même contrat de stockage, sauf si l'hôte empêche l'événement annulable `file-command`. Examinez les diagnostics plutôt que de promettre un export sans perte. Les aperçus de `.vsd` binaires ne sont pas des cibles d'édition OneDrive natives.

Cela s'applique à l'élément `ooxml-ui/visio` ainsi qu'à `visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`, `visio-svelte-viewer`, `visio-solid-viewer` et `visio-vanilla-viewer`. `visio-core` est la façade du modèle. La connexion n'y ajoute pas d'authentification Microsoft.

Voir la [référence Office OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#visio-vsdx) pour le câblage des paquets, l'authentification, le partage tenant compte des autorisations et la récupération en cas de conflit qui conserve les deux versions. Voir les [tests](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) pour les transports locaux aux tests et la validation sur le service réel. Aucune interface de simulation ni aucun service n'est livré.

Enregistrez `mergeDrawingSnapshots` de `ooxml-core/visio/collab` comme rappel `merge` de la session pour combiner, lors de l'enregistrement, les modifications indépendantes prises en charge sur les formes et les propriétés. La suite le fait automatiquement. Les modifications qui se chevauchent ou sont ambiguës exigent toujours une récupération. Il ne s'agit ni de la coédition Microsoft ni d'un verrou de fichier. Les [limites de Visio](/parity) existantes et les diagnostics d'export restent applicables.

## Ouvrir un fichier cloud depuis votre propre interface {#open-a-cloud-file-from-your-own-ui}

Fichier > Ouvrir n'est accessible qu'une fois un document ouvert. Pour ouvrir un fichier depuis une page d'accueil ou votre propre sélecteur, obtenez un `CloudOpenedFile` depuis n'importe quel `CloudFileProvider` (le navigateur partagé `<office-ui-cloud-files>` l'émet dans `office-cloud-file-open`) et transmettez-le à `viewer.openCloudFile(file)` sur `<visio-viewer>` (le handle de chaque liaison l'expose aussi). Cela suit le même chemin que Fichier > Ouvrir > OneDrive : le nom du fichier, la lecture seule lorsque `canEdit` vaut false, et l'enregistrement routé vers le stockage propre au fichier jusqu'au chargement d'un autre document.
