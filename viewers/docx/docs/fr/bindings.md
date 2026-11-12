# Adaptateurs de framework

Tous les adaptateurs montent `<docx-editor>` par la même fonction `mountEditor`. Une affectation de propriété est un remplacement externe du document, et non un événement de modification. `document-change` transporte directement un `DocumentModel` ; `document-error` transporte une `Error`. Les deux événements remontent à travers la frontière du shadow DOM.

React :

```tsx
import { WordEditor } from 'docx-react-viewer';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` expose `element`, `load(bytes)`, `save()`, `download()`, `markClean()`, `openCloudFile()` et `dirty`.

Vue :

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular :

```ts
import { WordEditorComponent } from 'docx-angular-viewer';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5 :

```svelte
<script>
  import WordEditor from 'docx-svelte-viewer';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

SolidJS :

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from 'docx-solid-viewer';

function DocumentEditor() {
	const [model, setModel] = createSignal(initialDocument);
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

Solid utilise les mêmes options `documentModel`, `readOnly`, `locale`, `onDocumentChange` et `onDocumentError` que l'adaptateur partagé. Utilisez `editorRef` pour obtenir les méthodes `element`, `load` et `save` de l'éditeur monté. Le propriétaire Solid nettoie l'éditeur et ses écouteurs lorsque le composant est démonté.

Élément personnalisé :

```ts
import { registerDocxEditor } from 'docx-vanilla-viewer';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

L'adaptateur de montage vanilla est exporté depuis `docx-vanilla-viewer` et renvoie les opérations update, destroy, load, save et element. Il évite les réinitialisations en boucle lorsqu'un parent renvoie le modèle émis. Fournissez un nouvel objet pour les changements de modèle externes ; modifier un modèle sur place n'est pas un mécanisme de réactivité pris en charge.

Les paquets de framework utilisent leur propre cycle de vie natif et leurs propres API de référence. Aucun modèle de framework ne contient de barre d'outils, de moteur de rendu des pages ni d'implémentation de commande de document. Les imports SSR n'enregistrent aucun élément ; le montage est une opération côté client. Le contrat actuel de navigateur valide les versions installées, et non chaque version historique des dépendances peer.

Définissez la propriété facultative `locale` sur n'importe quel éditeur de framework pour localiser le ruban et les contrôles de recherche (`en`, `fr`, `de`, `es` ou `zh-CN` ; par exemple, `locale="de"`). Cela n'affecte que le texte de l'interface ; cela ne traduit pas le contenu du document.

## Habillage de la fenêtre et commandes de fichier {#window-chrome-and-file-commands}

L'éditeur dessine, dans sa shadow root, un habillage de style Word : une barre de titre (Enregistrer, Annuler et Rétablir dans la barre d'accès rapide, le nom du fichier et l'état d'enregistrement, une recherche de commande « Dites-moi ce que vous voulez faire »), le sélecteur Édition/Affichage, **Commentaires** et **Partager** à l'extrémité droite de la rangée d'onglets du ruban (comme dans Word 365), un onglet **Fichier** avec un backstage (Informations avec les notes de compatibilité, Nouveau, Ouvrir, Enregistrer, Enregistrer une copie au format DOCX, Imprimer), et une barre d'état (page, mots, notes de compatibilité, affichage Web/Mise en page et zoom).

Définissez `element.fileName` pour nommer le document ouvert. Les commandes de fichier sont d'abord annoncées par un événement `file-command` annulable, dont `detail.command` vaut `new`, `open`, `save`, `saveAs`, `export`, `exportText`, `print` ou `share`. Appelez `preventDefault()` pour le traiter dans votre application (par exemple pour enregistrer dans votre propre stockage) ; sinon, l'éditeur utilise son comportement par défaut, réservé au navigateur : un sélecteur de fichiers, un téléchargement ou la boîte de dialogue d'impression. `share` (le bouton Partager) ouvre la boîte de dialogue de partage utilisée par tous les éditeurs Office : démarrez une session (nom de la session, votre nom et un serveur facultatif) ou rejoignez-en une à partir d'un lien d'invitation ou du nom de la session, puis copiez le lien d'invitation, consultez les participants et arrêtez le partage, ce qui définit la propriété `collaboration` ; annulez-la pour partager à votre façon. Le choix de **Affichage** déclenche `readonly-change`.

```ts
editor.fileName = 'Quarterly report.docx';
editor.addEventListener('file-command', (event) => {
	const { command } = (event as CustomEvent<{ command: string }>).detail;
	if (command === 'save') {
		event.preventDefault();
		void editor.save().then(uploadToMyStorage);
	}
});
```

## Enregistrement, état des modifications et options de l'interface {#saving-dirty-state-and-ui-options}

`save()` renvoie un `Blob` du document sérialisé (`saveBytes()` renvoie le `Uint8Array` brut), `download(fileName?)` enregistre puis lance un téléchargement du navigateur, et `dirty` / `dirty-change` signalent les modifications non enregistrées. `dirty` passe à vrai sur `document-change` et à faux après Fichier > Enregistrer, `download()`, le chargement d'un document ou `markClean()` ; `save()` seul ne le remet pas à faux, car seul l'hôte sait si le Blob a été conservé. Chaque adaptateur expose `save`, `download`, `markClean` et `dirty` sur son handle, ainsi que `onDirtyChange` / `onPageChange` (ou l'équivalent événementiel du framework).

| Propriété (attribut)                 | Par défaut | Effet                                                                                                                                                                                                                                                                                                                                                                                      |
| ------------------------------------ | ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `showThumbnails` (`show-thumbnails`) | `false`    | Colonne de gauche avec les miniatures des pages. Les miniatures nécessitent la mise en page Impression ; en mode Brouillon, la colonne l'indique au lieu de lister les pages.                                                                                                                                                                                                              |
| `showToolbar` (`show-toolbar`)       | `true`     | `show-toolbar="false"` masque le ruban ; la barre de titre et la barre d'état restent.                                                                                                                                                                                                                                                                                                     |
| `hiddenActions`                      | `[]`       | Contrôles du ruban à masquer, par identifiant stable en kebab-case (`RIBBON_ACTION_IDS`, par exemple `'bold'`, `'insert-table'`, `'track-changes'`, `'paste'`, `'word-count'`) ; les groupes et onglets vides sont masqués aussi. Un bouton de couleur et sa flèche de liste se masquent ensemble. Les identifiants inconnus sont ignorés. Les identifiants ne changent pas avec `locale`. |
| `ribbonAddIns`                       | aucun      | Onglets que l'hôte ajoute après les onglets intégrés ; voir [Onglets de complément du ruban](#ribbon-add-in-tabs).                                                                                                                                                                                                                                                                         |

**Déprécié :** avant l'existence des identifiants, `hiddenActions` prenait le libellé anglais du contrôle (`'Bold'`, `'Insert table'`). Ces chaînes `LegacyRibbonLabel` sont encore acceptées pendant une version : elles sont converties en identifiants, et l'élément émet un événement `document-warning` par affectation. La lecture de `hiddenActions` renvoie toujours des identifiants. Les événements `ribbon-action` ne transportaient jamais de libellés (leur détail est indexé par type de commande), ils ne changent donc pas.

`page-change` (`{ page, pageCount }`) n'est émis qu'en mise en page Impression. La pagination provient du moteur de mise en page de cet éditeur et reste une approximation, et non la pagination de Word ; la barre d'état et les info-bulles des miniatures l'indiquent.

## Onglets de complément du ruban {#ribbon-add-in-tabs}

Un hôte peut ajouter ses propres onglets de ruban après ceux de Word, comme le fait un complément Office. `ribbonAddIns` prend des données simples et est une propriété de chaque adaptateur ; une commande choisie exécute son callback `run` et est signalée par `onRibbonAddIn` (Vue `office-ribbon-add-in`, Angular `ribbonAddIn`, Svelte `onribbonaddin`) avec `{ tab, command }`.

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [
					{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() },
					{ id: 'options', label: 'Options', size: 'small' },
					{ id: 'send', label: 'Send', items: [{ id: 'send-mail', label: 'By mail' }] },
				],
			},
		],
	},
];

<WordEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

`size: 'small'` place les commandes en colonnes de trois et `items` crée une liste déroulante. Transmettre à nouveau un descripteur égal, comme le fait un framework à chaque rendu, conserve les panneaux et remplace seulement les callbacks ; omettre la propriété supprime les onglets. Un onglet ne peut pas prendre l'identifiant d'un onglet intégré, et ses commandes n'atteignent jamais la gestion de commandes propre à l'éditeur. Le même descripteur (`RibbonAddInTab`, de `ooxml-core/ribbon`, réexporté par chaque paquet de framework) fonctionne dans les éditeurs Word, Excel, PowerPoint et Visio. Aucun onglet de complément n'est livré avec l'éditeur.

## Propriétés et événements de l'adaptateur {#binding-props-and-events}

Chaque adaptateur de framework transmet les mêmes propriétés de l'élément et réémet les mêmes événements : les noms ci-dessous ont donc le même sens dans chaque adaptateur. Les noms d'événements sont ceux de l'élément (`file-command`) ; chaque framework les écrit selon sa propre syntaxe.

| Propriété de l'élément | React / Solid / Vanilla | Propriété Vue  | Entrée Angular | Propriété Svelte |
| ---------------------- | ----------------------- | -------------- | -------------- | ---------------- |
| `fileName`             | `fileName`              | `fileName`     | `fileName`     | `fileName`       |
| `reviewAuthor`         | `reviewAuthor`          | `reviewAuthor` | `reviewAuthor` | `reviewAuthor`   |
| `themeColors`          | `themeColors`           | `themeColors`  | `themeColors`  | `themeColors`    |

`fileName`, `reviewAuthor` et `themeColors` sont transmis dès que la propriété est définie ou change, de sorte qu'un renommage effectué par Fichier > Ouvrir ne soit pas annulé par un rendu qui transmet l'ancien nom.

| Événement de l'élément | Option React / Solid / Vanilla | Événement Vue      | Sortie Angular    | Callback Svelte     |
| ---------------------- | ------------------------------ | ------------------ | ----------------- | ------------------- |
| `file-command`         | `onFileCommand(event)`         | `file-command`     | `fileCommand`     | `onfilecommand`     |
| `document-warning`     | `onDocumentWarning(message)`   | `document-warning` | `documentWarning` | `ondocumentwarning` |
| `readonly-change`      | `onReadOnlyChange(readOnly)`   | `readonly-change`  | `readOnlyChange`  | `onreadonlychange`  |
| `ribbon-action`        | `onRibbonAction(detail)`       | `ribbon-action`    | `ribbonAction`    | `onribbonaction`    |
| `ribbon-customize`     | `onRibbonCustomize(ids)`       | `ribbon-customize` | `ribbonCustomize` | `onribboncustomize` |

`file-command` est l'unique événement annulable : son gestionnaire reçoit donc l'événement lui-même et non son détail. Appelez `event.preventDefault()` pour prendre en charge une commande Fichier (Enregistrer, Ouvrir, Nouveau, etc.) ; `event.detail.command` indique laquelle. Les gestionnaires Vue, Angular et Svelte reçoivent le même objet `CustomEvent`. Les autres événements transmettent directement leur détail.
