# Intégration OneDrive

Pour une configuration pas à pas (inscription Microsoft Entra, connexion et un exemple complet), commencez par le [démarrage rapide OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

Pour d'autres services ou des flux OneDrive personnalisés, utilisez la même propriété `storage` avec [`createStorageSession` et le guide d'intégration indépendant du fournisseur](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md). Les hôtes peuvent remplacer l'authentification, le transport, les gestionnaires de fusion, l'interface de conflit et la politique de partage sans modifier aucun adaptateur Word.

Tous les paquets Word utilisent la connexion partagée `ooxml-core/embedding/onedrive`. L'hôte résout le lien de partage, fournit l'authentification Microsoft et l'autorisation d'écriture vérifiée, et relie la session obtenue au contrat de stockage existant de l'éditeur.

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
editor.fileName = opened.file.name;
await editor.load(opened.bytes);
```

Le `<docx-editor>` simple de `ooxml-ui/docx` expose `saveBytes()`. Chaque handle de framework expose `save()` qui renvoie un `Blob`, `load(bytes)`, `markClean()` et `dirty`. Un bouton Enregistrer personnalisé peut convertir le Blob en octets et appeler `session.save(bytes)` ; ne marquez le document comme propre que lorsque la persistance a réussi et qu'aucune modification plus récente n'est intervenue. L'enregistrement natif passe automatiquement par `storage`, sauf si l'hôte empêche l'événement `file-command` annulable.

## Ouvrir depuis OneDrive dans Fichier > Ouvrir {#open-from-onedrive-in-file-open}

Définissez `fileProviders` (une propriété de l'élément et une prop de chaque adaptateur) avec une liste de `CloudFileProvider`s : Fichier > Ouvrir affiche alors des emplacements : « Cet appareil » (le bouton Parcourir) et un emplacement par fournisseur, chacun affichant le navigateur partagé `<office-ui-cloud-files>` (Récents, Mes fichiers, recherche, ouverture depuis un lien de partage) filtré sur `.docx`. Ses libellés proviennent des dictionnaires Word de chaque langue d'affichage.

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDocxSnapshots } from 'ooxml-core/docx';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your Microsoft authentication
		signIn, // opens your sign-in UI and resolves with the account
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			docx: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

Un fichier choisi dans un emplacement se charge comme n'importe quel document, prend son nom et passe en lecture seule lorsque le fournisseur indique `canEdit: false` (la plus stricte des deux valeurs, celle de l'hôte `readOnly` ou celle du fichier, l'emporte). Fichier > Enregistrer et Enregistrer sous vont vers le stockage propre à ce fichier au lieu du `storage` de l'hôte jusqu'au chargement d'un autre document ; la barre de titre n'indique « enregistré » qu'une fois l'envoi résolu, et un envoi échoué reste non enregistré et déclenche `document-error`. Ouvrir depuis un emplacement ne déclenche pas `file-command` (c'est l'hôte qui a choisi les fournisseurs) ; Enregistrer et Enregistrer sous le déclenchent toujours, et un hôte qui les annule les gère lui-même. Un hôte qui retire la capacité `open` masque la page et les emplacements avec elle.

La démo Word utilise par défaut `createDemoOneDriveProvider` de `demos/onedrive`. Connectez-vous depuis Fichier > Ouvrir > OneDrive : définissez `VITE_ONEDRIVE_CLIENT_ID` (et, facultativement, `VITE_ONEDRIVE_AUTHORITY`) dans `demos/docx/demo-vanilla/.env.local` pour utiliser votre propre inscription d'application Entra via MSAL, saisissez un identifiant client dans la boîte de dialogue de la démo, ou collez un jeton Graph Explorer dans son champ « Access token ». Enregistrez l'origine de la démo (avec une barre oblique finale) comme URI de redirection d'application monopage. La démo ne modifie que les fichiers du lecteur OneDrive de l'utilisateur connecté ; les autres fichiers s'ouvrent en lecture seule.

Cela s'applique à `docx-react-viewer`, `docx-vue-viewer`, `docx-angular-viewer`, `docx-svelte-viewer`, `docx-solid-viewer` et `docx-vanilla-viewer`. Chacun transmet les mêmes props/options `storage` et `readOnly`. `docx-core` est une façade de modèle et n'implémente pas la connexion Microsoft.

La [référence OneDrive d'Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#word-docx) contient la matrice des paquets, la configuration de l'authentification, les périmètres de partage et la récupération exploitable des conflits. Le [guide de test](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) sépare les transports locaux aux tests des essais d'acceptation Microsoft en conditions réelles. Aucune interface ni aucun service de simulation n'est livré.

Enregistrez `mergeDocxSnapshots` de `ooxml-core/docx` comme callback `merge` de la session pour combiner, lors de l'enregistrement, des modifications indépendantes de texte et de propriétés. La suite l'active automatiquement. Une structure ambiguë ou des remplacements incompatibles conservent les deux versions pour récupération. Cela ne relie pas les salles de collaboration Word à la co-écriture Microsoft Office. Les [limites de préservation](/fr/preservation) existantes s'appliquent toujours.

## Compte > Services connectés {#account-connected-services}

Les mêmes `fileProviders` alimentent aussi la page **Compte** de la vue Fichier avec une liste « Services connectés » (le composant partagé `<office-ui-connected-services>`) : une ligne par fournisseur, avec le compte connecté et les boutons Connecter ou Déconnecter, qui appellent les méthodes `signIn` et `signOut` du fournisseur. Un fournisseur sans `signOut` indique que la session est gérée par l'hôte. La liste suit le `subscribe` des fournisseurs : une connexion depuis Fichier > Ouvrir la met donc à jour, et ses textes proviennent des dictionnaires Word de chaque langue. Sans `fileProviders`, il n'y a pas de page Compte.

## Ouvrir un fichier cloud depuis votre propre interface {#open-a-cloud-file-from-your-own-ui}

Fichier > Ouvrir n'est accessible qu'une fois un document ouvert. Pour ouvrir un fichier depuis une page d'accueil ou votre propre sélecteur, obtenez un `CloudOpenedFile` auprès de n'importe quel `CloudFileProvider` (le navigateur partagé `<office-ui-cloud-files>` l'émet dans `office-cloud-file-open`), puis passez-le à `editor.openCloudFile(file)` sur `<docx-editor>` (le handle de chaque adaptateur l'expose aussi). Cela suit le même chemin que Fichier > Ouvrir > OneDrive : le nom du fichier, la lecture seule lorsque `canEdit` vaut faux, et l'enregistrement routé vers le stockage propre au fichier jusqu'au chargement d'un autre document.
