# Intégration OneDrive

Pour une configuration pas à pas (inscription Microsoft Entra, connexion et exemple complet), commencez par
le [démarrage rapide OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

## Fichier > Ouvrir > OneDrive {#file-open-onedrive}

Passez un fournisseur comme `fileProviders` et Fichier > Ouvrir affiche un emplacement OneDrive à côté de « Cet appareil » :

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your sign-in owns the token
		signIn,
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

L’emplacement affiche les fichiers récents et Mes fichiers, effectue des recherches et ouvre un lien de partage. Un classeur choisi
se charge comme n’importe quel autre, s’enregistre dans le même élément avec Enregistrer (une copie avec Enregistrer sous), et n’est
marqué comme enregistré qu’une fois le téléversement terminé ; un téléversement échoué le laisse non enregistré. Un fichier que `canWrite` refuse
(autre lecteur, fichier partagé) s’ouvre en lecture seule. `fileProviders` est une prop de chaque
liaison de framework (`fileProviders` dans React, Vue, Solid, Svelte, Angular et Vanilla). L’application de démonstration
(`demos/xlsx`) transmet par défaut le fournisseur OneDrive de démonstration ; configurez-le avec
`VITE_ONEDRIVE_CLIENT_ID` dans `demos/xlsx/demo-vanilla/.env.local`, ou collez un jeton de Graph Explorer dans sa boîte de dialogue de connexion. Le
branchement `storage` de plus bas niveau ci-dessous s’adresse aux hôtes qui construisent leur propre interface d’ouverture.

Le même contrat de stockage prend aussi en charge des fournisseurs personnalisés. Consultez
[la configuration du stockage et de la fusion indépendante des fournisseurs](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)
pour `createStorageSession`, les remplacements de fusion par format et les paramètres d’amorçage de la suite.
Utilisez ces points d’extension avec chaque paquet de framework Excel.

Excel utilise la connexion commune `ooxml-core/embedding/onedrive`. Résolvez le lien de partage dans
l’hôte, vérifiez les autorisations propres à l’élément et donnez à l’éditeur existant les octets XLSX téléchargés :

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
await editor.load(opened.bytes, opened.file.name);
```

L’élément et les poignées des frameworks exposent `saveBytes('xlsx')`, `load(bytes, fileName?)` et
`markClean()`. Un enregistrement personnalisé de l’hôte appelle `session.save(await editor.saveBytes('xlsx'))` ; ne marquez
comme enregistré qu’une fois la persistance réussie et sans modifications plus récentes. Utilisez le chemin d’enregistrement normal de l’éditeur pour valider
une modification de cellule active. Le CSV est un export distinct et ne doit jamais remplacer l’élément XLSX connecté.
L’enregistrement natif passe par `storage`, sauf si l’hôte annule l’événement `file-command` annulable.

Le même contrat `storage` et `readOnly` s’applique à `@christophervr/xlsx-react-viewer`,
`xlsx-vue-viewer`, `xlsx-angular-viewer`, `xlsx-svelte-viewer`, `xlsx-solid-viewer` et
`xlsx-vanilla-viewer`, ainsi qu’à l’élément simple `ooxml-ui/xlsx`. `@christophervr/xlsx-core`
est une façade du modèle, et non un service d’authentification.

Consultez la [référence du paquet OneDrive Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#excel-xlsx)
pour l’authentification, le partage et la récupération qui conserve à la fois les versions locale et distante. Le
[guide de test](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) explique les jeux de tests sans identifiants et la
validation sur un compte réel ; aucune interface ni aucun service de simulation n’est livré.

Enregistrez `mergeWorkbookSnapshots` de `ooxml-core/xlsx/collab` comme rappel `merge` de la session
pour combiner, lors de l’enregistrement, les modifications indépendantes et prises en charge d’un classeur. La suite l’active
automatiquement ; les modifications de cellules qui se chevauchent et les changements non pris en charge exigent toujours une récupération. Cela ne connecte pas
les sessions de collaboration sur classeur à la coédition de Microsoft Excel, n’actualise pas les données externes et ne modifie pas les
[limites fonctionnelles](/fr/features).

## Compte > Services connectés {#account-connected-services}

Les mêmes `fileProviders` remplissent aussi la page **Compte** de la vue Fichier avec une liste
« Services connectés » (le `<office-ui-connected-services>` partagé) : une ligne par fournisseur avec le compte connecté et
Se connecter ou Se déconnecter, qui appellent les `signIn` et `signOut` propres au fournisseur. Un fournisseur sans
`signOut` indique que la session est gérée par l’hôte. La liste suit le `subscribe` des fournisseurs, si bien que
la connexion depuis Fichier > Ouvrir la met à jour, et ses textes proviennent des dictionnaires Excel dans toutes les
langues. Sans `fileProviders`, il n’existe pas de page Compte.

## Ouvrir un fichier cloud depuis votre propre interface {#open-a-cloud-file-from-your-own-ui}

Fichier > Ouvrir n’est accessible qu’une fois un document ouvert. Pour ouvrir un fichier depuis une page d’accueil ou votre propre
sélecteur, obtenez un `CloudOpenedFile` depuis n’importe quel `CloudFileProvider` (le navigateur partagé `<office-ui-cloud-files>`
l’émet dans `office-cloud-file-open`) et passez-le à `editor.openCloudFile(file)` sur `<xlsx-editor>` (la poignée de chaque liaison l’expose aussi). Il suit le
même chemin que Fichier > Ouvrir > OneDrive : le nom du fichier, la lecture seule lorsque `canEdit` vaut false, et Enregistrer
dirigé vers le stockage propre au fichier jusqu’au chargement d’un autre document.
