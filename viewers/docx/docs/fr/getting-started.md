# Prise en main

docx-viewer se compose d'un composant web `<docx-editor>` et de six fines couches d'adaptation pour les frameworks. Cette page installe un adaptateur, monte l'éditeur, puis charge et enregistre un document. Il s'agit d'une implémentation précoce : elle n'offre pas la parité avec Microsoft Word ni un export sans perte (voir la [feuille de route de la prise en charge](/parity-roadmap)).

## 1. Installer

Installez le paquet autonome correspondant à votre framework. Il embarque l'interface de l'éditeur et l'adaptateur, et apporte `docx-core` (le modèle de document) avec lui.

::: code-group

```bash [React]
npm install docx-react-viewer react
```

```bash [Vue]
npm install docx-vue-viewer vue
```

```bash [Angular]
npm install docx-angular-viewer
```

```bash [Svelte]
npm install docx-svelte-viewer svelte
```

```bash [Solid]
npm install docx-solid-viewer solid-js
```

```bash [Vanilla JS]
npm install docx-vanilla-viewer
```

:::

N'installez `docx-core` seul que pour un usage sans interface (analyser et sérialiser sans éditeur).

## 2. Monter l'éditeur

Chaque adaptateur monte le même élément `<docx-editor>`. Transmettez un `DocumentModel`, écoutez les changements et donnez une hauteur au conteneur.

```tsx
import { useState } from 'react';
import { createDocument, WordEditor } from 'docx-react-viewer';

export function Editor() {
	const [model, setModel] = useState(() => createDocument());
	return (
		<div style={{ height: '100vh' }}>
			<WordEditor documentModel={model} onDocumentChange={setModel} />
		</div>
	);
}
```

Les autres frameworks utilisent les mêmes options : consultez les [guides par framework](/frameworks/react) et le [contrat des adaptateurs](/bindings).

## 3. Charger et enregistrer des fichiers

La référence du composant (ou le handle vanilla) expose `load(bytes)` et `save()`. Le chargement détecte le DOCX et l'ancien format `.doc` ; les fichiers protégés par mot de passe ne sont pas pris en charge.

```ts
const bytes = new Uint8Array(await file.arrayBuffer());
await editor.load(bytes);

// After edits:
const blob = await editor.save(); // Blob of the serialized .docx
```

`dirty` et l'événement `dirty-change` signalent les modifications non enregistrées. Consultez l'[API de l'élément](/fr/api) pour toutes les propriétés, méthodes et tous les événements.

## 4. L'essayer

- Les [démos en direct](/demos) exécutent le même éditeur dans chaque adaptateur de framework.
- Pour lancer la démo depuis un clone du dépôt (Bun requis) :

```bash
bun install
bun run demo
```

## Étapes suivantes

- [API de l'élément et événements](/fr/api)
- [Thèmes et synchronisation clair/sombre](/theming)
- [Édition du texte](/fr/editing) et [collaboration](/fr/collaboration)
- [Architecture](/fr/architecture) et [unités du modèle](/model-units)
