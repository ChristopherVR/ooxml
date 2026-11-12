---
title: Intégration OneDrive
description: Connecter chaque binding de visionneuse PowerPoint au stockage partagé Office OneDrive, au partage et à la récupération après conflit.
---

# Intégration OneDrive

Chaque package de framework PowerPoint accepte aussi un fournisseur de stockage personnalisé via sa prop `storage` existante. Voir [intégration indépendante du fournisseur et configuration](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md) pour `createStorageSession`, les surcharges d'authentification et de transport, les gestionnaires de fusion par format, les boîtes de dialogue de conflit de l'hôte et les choix de partage explicites.

PowerPoint utilise la même connexion facultative `ooxml-core/embedding/onedrive` que Word, Excel et Visio. L'hôte fournit l'authentification Microsoft et les autorisations d'éléments vérifiées. Il n'existe ni enregistrement d'application Microsoft intégré, ni service de simulation, ni démonstration OneDrive simulée.

Pour une configuration pas à pas avec captures d'écran (enregistrement de l'application, connexion, case de lien de partage et un exemple React complet), commencez par le [démarrage rapide OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md). Chaque démonstration de framework propose OneDrive sous Fichier > Ouvrir pour l'essayer avec votre propre OneDrive ; la section « Essayer dans les démonstrations PowerPoint » du démarrage rapide explique la configuration.

## Emplacements de Fichier > Ouvrir : `fileProviders` {#file-open-places-fileproviders}

Passez `fileProviders` (une liste de `CloudFileProvider`, par exemple issue de `createOneDriveFileProvider` dans `ooxml-core/embedding/onedrive`) et la page Ouvrir affiche **Cet appareil** ainsi qu'un emplacement par fournisseur. La prop s'écrit de la même façon dans chaque binding : `fileProviders` (React, entrée Angular, Svelte, option vanilla) et `:file-providers` (Vue).

```tsx
<PowerPointViewer content={bytes} canEdit fileProviders={[oneDriveProvider]} />
```

Pour combiner les modifications qu'une autre personne a enregistrées entre-temps, donnez au fournisseur un fusionneur PowerPoint. Sans lui, une version plus récente sur OneDrive ouvre toujours les choix de conflit ; `merge: false` demande la même chose pour chaque format, et `maxMergeAttempts` (3 par défaut lorsqu'une fusion s'applique) borne le nombre de nouvelles tentatives :

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergePptxSnapshots } from 'ooxml-core/pptx';

const oneDriveProvider = createOneDriveFileProvider({
	getAccessToken,
	canWrite,
	signIn,
	mergers: {
		pptx: ({ base, local, remote }) => mergePptxSnapshots(base.bytes, local, remote.bytes),
	},
});
```

Un fichier ouvert depuis un emplacement se charge comme n'importe quelle ouverture (historique d'annulation réinitialisé, `onLoad` se déclenche). La visionneuse utilise ensuite le stockage propre à ce fichier pour Fichier > Enregistrer, Enregistrer sous et les copies de récupération, au lieu de la prop `storage`, et l'ouvre en lecture seule lorsque le service n'autorise pas la modification ; un hôte `canEdit={false}` ou `readOnly` reste en lecture seule. Cela dure jusqu'au chargement d'un autre document : une nouvelle présentation, une ouverture depuis cet appareil ou depuis Récents, ou un nouveau `content` fourni par l'hôte. Un échec d'enregistrement affiche la notification d'échec et la barre d'état conserve **Modifications non enregistrées**.

## Compte > Services connectés {#account-connected-services}

Les mêmes `fileProviders` alimentent aussi une section **Services connectés** de **Fichier > Compte**, sous les cartes profil, stockage et connexion. Chaque fournisseur est une ligne avec son compte (nom et e-mail) et les boutons **Connecter** ou **Déconnecter**, qui appellent respectivement les méthodes `signIn` et `signOut` du fournisseur ; un fournisseur sans `signOut` indique que sa session est gérée par votre organisation. La liste suit `subscribe` du fournisseur : une connexion depuis Fichier > Ouvrir met donc la liste à jour. Sans `fileProviders`, rien de supplémentaire n'apparaît, et la carte de connexion existante `accountAuth` ne change pas. Les cinq bindings affichent l'élément partagé `office-ui-connected-services` (sous le nom `pptx-ui-connected-services`) avec des libellés traduits.

Le [guide d'intégration Office complet](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) documente l'authentification, le stockage des sessions, les copies cloud sûres, les rôles et portées de partage, la récupération après conflit et la référence de chaque type de package.

```ts
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';

const drive = createOneDriveClient({
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
});
const opened = await drive.open(sharingLink);
const session = createOneDriveSession(drive, opened);
// Pass session.storage and opened.file.access === 'write' to the viewer.
```

La session et les services d'autorisation de cet exemple sont des intégrations de l'hôte, et non des objets fournis. Importez l'adaptateur comme dépendance directe de `ooxml-core`. `ooxml-ui/pptx/onedrive` ne reste qu'une surface de compatibilité pour l'ancien import propre à PowerPoint.

## Les cinq bindings {#all-five-bindings}

| Binding/package                | Ouvrir les octets                     | Restreindre la modification   | Exporter le PPTX actuel           |
| ------------------------------ | ------------------------------------- | ----------------------------- | --------------------------------- |
| React, `pptx-react-viewer`     | `content={opened.bytes}`              | `canEdit={canEdit}`           | `await ref.current.getContent()`  |
| Vue, `pptx-vue-viewer`         | `:content="opened.bytes"`             | `:can-edit="canEdit"`         | `await viewer.value.getContent()` |
| Angular, `pptx-angular-viewer` | `[content]="opened.bytes"`            | `[canEdit]="canEdit"`         | `await viewer.getContent()`       |
| Svelte, `pptx-svelte-viewer`   | `source={opened.bytes}`               | `editable={canEdit}`          | `await viewer.getContent()`       |
| Vanilla, `pptx-vanilla-viewer` | `await viewer.loadFile(opened.bytes)` | `viewer.setEditable(canEdit)` | `await viewer.save()`             |

`canEdit` (`editable` dans Svelte) signifie `opened.file.access === 'write'`. Un accès inconnu reste en lecture seule. Chaque binding accepte `storage` comme prop, entrée ou option existante. Affectez `session.storage` pour connecter l'enregistrement natif de Fichier à OneDrive. `onOpenFile` peut ouvrir le sélecteur de liens de partage de l'hôte ; ne passez pas l'URL de la page de partage Microsoft à un chargeur de source de document générique.

La visionneuse ne marque la présentation comme enregistrée qu'après la résolution de `storage.save`, et uniquement si rien n'a été modifié pendant l'envoi ; un enregistrement rejeté laisse la présentation non enregistrée. Pour un bouton Enregistrer géré par l'hôte, exportez le PPTX actuel et appelez `session.save(bytes)`. Gardez l'éditeur monté et préservez l'historique d'annulation. Suivez séparément les modifications faites pendant l'envoi, par rapport à l'instantané envoyé. Ne considérez pas une méthode d'enregistrement native résolue comme une preuve de réussite : certaines méthodes signalent les échecs via `onError` et la pile de notifications. Voir [stockage géré par l'hôte](/fr/guide/host-owned-storage).

## Conflits et partage {#conflicts-and-sharing}

Une modification de révision seule peut être retentée une fois lorsque les octets téléchargés depuis le cloud sont toujours égaux à la référence chargée. Un véritable conflit de contenu conserve les versions de référence, locale et distante dans `OneDriveConflictError`. Proposez **Enregistrer mes modifications comme copie OneDrive**, **Garder la copie locale et ouvrir la dernière version**, **Télécharger mes modifications** et **Continuer la modification**. N'écrasez jamais silencieusement et ne rechargez jamais sans avertir. Enregistrez `mergePptxSnapshots` comme callback `merge` de la session pour combiner automatiquement les modifications indépendantes et prises en charge des diapositives. La suite le fait par défaut. Les chevauchements non pris en charge exigent toujours une récupération ; la session ne verrouille pas le fichier jusqu'à la validation finale de l'envoi.

Enregistrer sous vers un PPTX natif compatible crée une copie cloud distincte via le stockage de la session. Les exports vers PPT/PPTM/PPSX ou un autre format ne doivent jamais remplacer l'original PPTX : conservez-les comme un flux d'export distinct et explicite. Une opération de partage ne crée un lien qu'après que l'utilisateur a choisi son rôle d'affichage ou de modification et sa portée. Partager un lien OneDrive ne connecte pas une salle de collaboration PowerPoint à la coédition Microsoft PowerPoint.

Voir [tests d'intégration](/fr/guide/onedrive-testing) et les [limites de format](/fr/guide/limitations) existantes.
