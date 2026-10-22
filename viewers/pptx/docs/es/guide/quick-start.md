---
title: Inicio rápido
description: Cuatro flujos de extremo a extremo - crear una presentación, analizar y editar una existente, convertir a Markdown y renderizar con el visualizador de React.
---

# Inicio rápido

Esta página recorre cuatro flujos comunes de extremo a extremo. Cada uno es autónomo y usa la API pública real. Para referencias más detalladas, siga los enlaces de [Siguientes pasos](#siguientes-pasos).

## 1. Crear una presentación desde cero

Use `PptxHandler.create()` para iniciar una presentación nueva, construya diapositivas con el constructor de diapositivas fluido y luego llame a `save()` para obtener los bytes.

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

## 2. Analizar y editar una presentación existente

Construya un `PptxHandler`, llame a `load()` con un `ArrayBuffer`, recorra el [modelo de datos](/es/guide/data-model), modifíquelo y llame a `save()`.

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

::: tip Estrechar los elementos
`slide.elements` es un arreglo de la unión discriminada [`PptxElement`](/es/guide/data-model). Compruebe siempre `element.type` antes de acceder a campos específicos de cada variante.
:::

## 3. Convertir a Markdown

`PptxMarkdownConverter` convierte el `PptxData` analizado en Markdown, extrayendo opcionalmente los medios a una carpeta.

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

## 4. Renderizar con el visualizador de React

El componente `PowerPointViewer` renderiza una presentación a partir de un `Uint8Array` y puede ejecutarse en modo editable. La devolución de llamada `onContentChange` recibe los bytes `.pptx` reserializados cada vez que se edita la presentación.

![The editable viewer with ribbon, thumbnails, and inspector](/docs-shots/editor.jpg)

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
Consulte [Props del componente](/react/props) para la referencia completa de props verificada contra el código fuente: el visualizador recibe contenido `Uint8Array` (no `ArrayBuffer`), y `onContentChange` entrega bytes serializados en lugar de un indicador de cambios.
:::

¿Usa Vue, Angular, Svelte o ningún framework? Consulte las guías de primeros pasos de [Vue](/vue/getting-started), [Angular](/angular/getting-started), [Svelte](/svelte/getting-started) o [Vanilla JS](/vanilla/getting-started) para el recorrido equivalente en cada enlace.

## Siguientes pasos

- [Resumen del paquete Core](/core/) - las APIs completas de handler, constructor y conversor.
- [Resumen del paquete React](/react/) - props del visualizador, edición, presentación y exportación.
- [El modelo PptxData](/es/guide/data-model) - la forma de las presentaciones analizadas.
- [El modelo PptxData](/es/guide/data-model#unidades-emu-y-pixeles) - unidades EMU, el modelo de elementos y la resolución del tema.
