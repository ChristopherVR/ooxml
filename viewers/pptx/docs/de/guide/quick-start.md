---
title: Schnellstart
description: Vier vollständige Abläufe - eine Präsentation erstellen, eine analysieren und bearbeiten, in Markdown konvertieren und mit dem React-Viewer rendern.
---

# Schnellstart

Diese Seite führt durch vier häufige Abläufe von Anfang bis Ende. Jeder ist eigenständig und nutzt die echte öffentliche API. Tiefergehende Referenzen finden Sie über die Links unter [Nächste Schritte](#nachste-schritte).

## 1. Eine Präsentation von Grund auf erstellen

Verwenden Sie `PptxHandler.create()`, um eine neue Präsentation zu starten, bauen Sie Folien mit dem fließenden Folien-Builder und rufen Sie dann `save()` auf, um Bytes zu erhalten.

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

## 2. Eine vorhandene Präsentation analysieren und bearbeiten

Erzeugen Sie einen `PptxHandler`, rufen Sie `load()` mit einem `ArrayBuffer` auf, durchlaufen Sie das [Datenmodell](/de/guide/data-model), verändern Sie es und rufen Sie `save()` auf.

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

::: tip Elemente eingrenzen
`slide.elements` ist ein Array der diskriminierten Union [`PptxElement`](/de/guide/data-model). Prüfen Sie immer `element.type`, bevor Sie auf variantenspezifische Felder zugreifen.
:::

## 3. In Markdown konvertieren

`PptxMarkdownConverter` wandelt analysierte `PptxData` in Markdown um und extrahiert auf Wunsch Medien in einen Ordner.

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

## 4. Mit dem React-Viewer rendern

Die Komponente `PowerPointViewer` rendert eine Präsentation aus einem `Uint8Array` und kann im Bearbeitungsmodus laufen. Der Callback `onContentChange` erhält die neu serialisierten `.pptx`-Bytes, sobald die Präsentation bearbeitet wird.

![Der bearbeitbare Viewer mit Menüband, Miniaturansichten und Inspektor](/docs-shots/editor.jpg)

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
Die vollständige, am Quellcode geprüfte Prop-Referenz steht unter [Komponenten-Props](/react/props) - der Viewer erwartet `Uint8Array`-Inhalt (nicht `ArrayBuffer`), und `onContentChange` liefert serialisierte Bytes statt eines Dirty-Flags.
:::

Sie verwenden Vue, Angular, Svelte oder gar kein Framework? Die Einstiegsanleitungen für [Vue](/vue/getting-started), [Angular](/angular/getting-started), [Svelte](/svelte/getting-started) oder [Vanilla JS](/vanilla/getting-started) zeigen die entsprechende Schrittfolge für jedes Binding.

## Nächste Schritte

- [Core-Paketübersicht](/core/) - die vollständigen APIs für Handler, Builder und Konverter.
- [React-Paketübersicht](/react/) - Viewer-Props, Bearbeiten, Präsentieren und Export.
- [Das PptxData-Modell](/de/guide/data-model) - die Struktur analysierter Präsentationen.
- [Das PptxData-Modell](/de/guide/data-model#einheiten-emu-und-pixel) - EMU-Einheiten, das Elementmodell und die Designauflösung.
