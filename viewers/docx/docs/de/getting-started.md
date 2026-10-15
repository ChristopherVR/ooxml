# Erste Schritte

docx-viewer besteht aus einer Web Component `<docx-editor>` und sechs schlanken Framework-Adaptern. Diese Seite zeigt, wie Sie einen Adapter installieren, den Editor einbinden sowie ein Dokument laden und speichern. Es handelt sich um eine frühe Implementierung: Sie bietet weder Parität mit Microsoft Word noch einen verlustfreien Export (siehe die [Roadmap zur Unterstützung](/parity-roadmap)).

## 1. Installation

Installieren Sie das eigenständige Paket für Ihr Framework. Es bündelt die Editor-Oberfläche und den Adapter und bringt `docx-core` (das Dokumentmodell) mit.

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

Installieren Sie `docx-core` nur dann separat, wenn Sie es headless verwenden (Parsen und Serialisieren ohne Editor).

## 2. Den Editor einbinden

Jeder Adapter bindet dasselbe Element `<docx-editor>` ein. Übergeben Sie ein `DocumentModel`, hören Sie auf Änderungen und geben Sie dem Container eine Höhe.

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

Die anderen Frameworks verwenden dieselben Optionen: siehe die [Framework-Leitfäden](/frameworks/react) und den [Binding-Vertrag](/bindings).

## 3. Dateien laden und speichern

Die Komponenten-Referenz (oder das Vanilla-Handle) stellt `load(bytes)` und `save()` bereit. Beim Laden werden DOCX und das ältere `.doc` erkannt; passwortgeschützte Dateien werden nicht unterstützt.

```ts
const bytes = new Uint8Array(await file.arrayBuffer());
await editor.load(bytes);

// After edits:
const blob = await editor.save(); // Blob of the serialized .docx
```

`dirty` und das Ereignis `dirty-change` melden ungespeicherte Änderungen. Alle Eigenschaften, Methoden und Ereignisse finden Sie in der [Element-API](/de/api).

## 4. Ausprobieren

- Die [Live-Demos](/demos) führen denselben Editor in jedem Framework-Adapter aus.
- So starten Sie die Demo aus einem Klon des Repositorys (Bun erforderlich):

```bash
bun install
bun run demo
```

## Nächste Schritte

- [Element-API und Ereignisse](/de/api)
- [Theming und Synchronisierung von Hell/Dunkel](/theming)
- [Text bearbeiten](/de/editing) und [Zusammenarbeit](/de/collaboration)
- [Architektur](/de/architecture) und [Modelleinheiten](/model-units)
