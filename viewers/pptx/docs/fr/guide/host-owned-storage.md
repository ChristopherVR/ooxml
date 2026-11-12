# Stockage géré par l'hôte

Pour un fournisseur Microsoft Graph de fichiers existants avec une démonstration sans enregistrement d'application, voir [Intégration OneDrive](/fr/guide/onedrive) et [tester OneDrive](/fr/guide/onedrive-testing).

Chaque binding accepte une prop `storage` (Angular : entrée ; Vanilla : option) qui confie les entrées-sorties de fichiers de la visionneuse à l'hôte. Il s'agit du contrat `DocumentStorage` de `ooxml-core/embedding`, le même qu'acceptent les éditeurs Word, Excel et Visio, et il est réexporté comme type depuis `ooxml-ui/pptx`.

```ts
import type { DocumentStorage } from 'ooxml-ui/pptx';

const storage: DocumentStorage = {
	// Required: File > Save and Save As.
	async save(bytes, { fileName, mimeType, reason }) {
		await uploadToMyServer(fileName, bytes, mimeType);
		// Optional: report the name the document now has.
		return { fileName };
	},
	// Optional, all three together: AutoRecover copies.
	readRecovery: (id) => myRecovery.get(id),
	writeRecovery: (entry) => myRecovery.put(entry),
	clearRecovery: (id) => myRecovery.delete(id),
	// Optional: lets Version History list the kept copies.
	listRecovery: () => myRecovery.list(),
};
```

## Enregistrer et Enregistrer sous {#save-and-save-as}

Sans `storage`, la visionneuse télécharge le fichier enregistré. Avec `storage`, Enregistrer et chaque format d'Enregistrer sous (`.pptx`, `.ppsx`, `.pptm`, `.ppt`) appellent `storage.save(bytes, meta)` à la place :

- `meta.fileName` est le nom de base de la source suivi de l'extension du format écrit (`savedPresentationFileName`) : `report.ppt` est donc enregistré sous `report.pptx`.
- `meta.mimeType` suit le format (`.ppt` est du type OLE2, `.ppsx` et `.pptm` ont chacun leur propre type).
- `meta.reason` est la commande exécutée : `'save'` pour Fichier > Enregistrer (la barre de titre, l'élément Enregistrer de la vue Backstage, la barre d'outils mobile et `downloadPptx()` de Vanilla/Svelte), `'saveAs'` pour chaque entrée d'Enregistrer sous, y compris Enregistrer sous > Présentation PowerPoint (.pptx).

Un échec d'enregistrement (la méthode `save()` de l'hôte a rejeté la promesse, ou la sérialisation a échoué) est signalé de la même façon dans les cinq bindings : l'événement `error` du binding (`onError`, `@error`, `(error)`, `onerror`) reçoit `Save failed: <reason>`, l'utilisateur voit la notification « Save failed » dans la pile de notifications de la visionneuse, et la copie de récupération automatique est conservée (elle n'est supprimée qu'après un enregistrement réussi et si Options > Enregistrer le demande). Les méthodes d'enregistrement résolvent la promesse au lieu de la rejeter.

Les octets sont les mêmes que ceux que renvoie `getContent()`, donc une présentation protégée par mot de passe est transmise chiffrée.

## Copies de récupération {#recovery-copies}

L'enregistrement automatique de récupération (les props `autosave` et `autosaveIntervalMs`, indexées par `filePath`) écrit normalement dans la base IndexedDB propre à la bibliothèque. Lorsque `storage` fournit `readRecovery`, `writeRecovery` et `clearRecovery`, chaque chemin de récupération les utilise à la place : l'écrivain de sauvegarde automatique, l'invite de récupération affichée au chargement (et son bouton Abandonner), le nettoyage après un enregistrement réussi, l'Historique des versions, Fichier > Ouvrir > Récents, Compte > Stockage et confidentialité et `restoreSessionDeck(storage)`. L'identifiant d'entrée est la clé `filePath` ; `savedAt` est un horodatage en millisecondes. Un stockage qui ne fournit qu'une partie des trois méthodes est ignoré pour la récupération, et IndexedDB reste utilisé.

La liste (Récents, utilisation du compte, Historique des versions) nécessite `listRecovery` ; sans elle, ces listes sont vides. Les opérations de maintenance d'Options > Enregistrer (« Supprimer les fichiers en cache », « vider le cache à la fermeture » et le balayage de « durée de conservation du cache ») ne suppriment les copies de l'hôte que si celui-ci fournit `listRecovery`. Sans elle, la visionneuse laisse les copies de l'hôte intactes, et la conservation relève du travail de l'hôte.

## Pas d'événement `file-command` {#no-file-command-event}

Les éléments Word, Excel et Visio annoncent un événement `file-command` annulable avant Enregistrer, Ouvrir et Nouveau. Les bindings PowerPoint n'ont pas un tel événement : pour prendre en charge l'enregistrement, passez `storage` ; pour prendre en charge Fichier > Ouvrir, utilisez `onOpenFile` (Vue et Angular : l'entrée `onOpenFile`).
