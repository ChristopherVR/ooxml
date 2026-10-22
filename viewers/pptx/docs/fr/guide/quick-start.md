---
title: Démarrage rapide
description: Quatre flux de bout en bout - créer une présentation, en analyser et en éditer une, convertir en Markdown et afficher avec la visionneuse React.
---

# Démarrage rapide

Cette page parcourt quatre flux courants de bout en bout. Chacun est autonome et utilise l'API publique réelle. Pour des références plus détaillées, suivez les liens de [Étapes suivantes](#etapes-suivantes).

## 1. Créer une présentation de toutes pièces

Utilisez `PptxHandler.create()` pour démarrer une nouvelle présentation, construisez des diapositives avec le constructeur de diapositives fluide, puis `save()` pour obtenir les octets.

```ts
import { writeFile } from 'node:fs/promises';

import { PptxHandler } from 'pptx-viewer-core';

const { handler, data, createSlide } = await PptxHandler.create({
	title: 'My Presentation',
	creator: 'Author Name',
	theme: {
		name: 'Custom Theme',
		colors: { accent1: '4472C4', accent2: 'ED7D31' },
		fonts: { majorFont: 'Calibri Light', minorFont: 'Calibri' },
	},
});

// Build a slide with the fluent API
const slide = createSlide()
	.addText('Hello World', { x: 100, y: 100, width: 600, height: 80, fontSize: 36 })
	.addShape('rect', { x: 100, y: 250, width: 300, height: 200 })
	.addImage('https://example.com/photo.jpg', { x: 450, y: 250, width: 300, height: 200 })
	.build();

data.slides.push(slide);

// Save to .pptx (returns a Uint8Array)
const output = await handler.save(data.slides);
await writeFile('presentation.pptx', output);
```

## 2. Analyser et éditer une présentation existante

Construisez un `PptxHandler`, appelez `load()` avec un `ArrayBuffer`, parcourez le [modèle de données](/fr/guide/data-model), modifiez-le, puis `save()`.

```ts
import { readFile, writeFile } from 'node:fs/promises';

import { PptxHandler } from 'pptx-viewer-core';

const handler = new PptxHandler();
const buffer = await readFile('presentation.pptx');
const data = await handler.load(
	buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
);

console.log(`Loaded ${data.slides.length} slides`);
console.log(`Theme: ${data.theme?.name}`);

// Access slide content - narrow on the `type` discriminant
for (const slide of data.slides) {
	for (const element of slide.elements) {
		if (element.type === 'text') {
			console.log(`Text: ${element.text}`);
		}
	}
}

// Modify and save
data.slides[0].elements[0].text = 'Updated Title';
const output = await handler.save(data.slides);
await writeFile('output.pptx', output);
```

::: tip Restreindre les éléments
`slide.elements` est un tableau de l'union discriminée [`PptxElement`](/fr/guide/data-model). Vérifiez toujours `element.type` avant d'accéder aux champs propres à une variante.
:::

## 3. Convertir en Markdown

`PptxMarkdownConverter` transforme les `PptxData` analysées en Markdown, avec une extraction optionnelle des médias dans un dossier.

```ts
import { readFile } from 'node:fs/promises';

import { PptxHandler, PptxMarkdownConverter } from 'pptx-viewer-core';

const handler = new PptxHandler();
const buffer = await readFile('presentation.pptx');
const data = await handler.load(
	buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
);

const converter = new PptxMarkdownConverter('./output', {
	sourceName: 'presentation.pptx',
	includeSpeakerNotes: true,
	mediaFolderName: 'media',
	includeMetadata: true,
	semanticMode: true, // Clean markdown vs positioned HTML
});

const markdown = await converter.convert(data);
console.log(markdown);
```

## 4. Afficher avec la visionneuse React

Le composant `PowerPointViewer` affiche une présentation à partir d'un `Uint8Array` et peut fonctionner en mode éditable. Le rappel `onContentChange` reçoit les octets `.pptx` resérialisés à chaque modification de la présentation.

![La visionneuse éditable avec ruban, miniatures et inspecteur](/docs-shots/editor.jpg)

```tsx
import { PowerPointViewer } from 'pptx-react-viewer/viewer';
import { useState } from 'react';

function App() {
	const [content, setContent] = useState<Uint8Array>();

	if (!content) return null;

	return (
		<PowerPointViewer
			content={content}
			canEdit
			onContentChange={(bytes) => {
				// `bytes` is the updated .pptx as a Uint8Array
				setContent(bytes);
			}}
		/>
	);
}
```

::: tip
Voir [Props des composants](/react/props) pour la référence complète des props, vérifiée dans le code source : la visionneuse prend un contenu `Uint8Array` (pas `ArrayBuffer`), et `onContentChange` fournit les octets sérialisés plutôt qu'un indicateur de modification.
:::

Vous utilisez Vue, Angular, Svelte ou aucun framework ? Consultez les guides de démarrage [Vue](/vue/getting-started), [Angular](/angular/getting-started), [Svelte](/svelte/getting-started) ou [Vanilla JS](/vanilla/getting-started) pour le parcours équivalent dans chaque binding.

## Étapes suivantes

- [Aperçu du package Core](/core/) - les API complètes de handler, builder et convertisseur.
- [Aperçu du package React](/react/) - props de la visionneuse, édition, présentation et export.
- [Le modèle PptxData](/fr/guide/data-model) - la forme des présentations analysées.
- [Le modèle PptxData](/fr/guide/data-model#unites-emu-et-pixels) - unités EMU, modèle d'éléments et résolution des thèmes.
