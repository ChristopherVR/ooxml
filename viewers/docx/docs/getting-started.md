# Getting started

docx-viewer is one `<docx-editor>` web component plus six thin framework adapters. This page installs an adapter, mounts the editor and loads and saves a document. It is an early implementation: it is not Microsoft Word parity and not lossless export (see the [support roadmap](/parity-roadmap)).

## 1. Install

Install the self-contained package for your framework. It bundles the editor UI and adapter and brings `docx-core` (the document model) with it.

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

Install `docx-core` on its own only for headless use (parse and serialize without an editor).

## 2. Mount the editor

Every adapter mounts the same `<docx-editor>` element. Pass a `DocumentModel`, listen for changes and give the container a height.

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

The other frameworks use the same options: see the [framework guides](/frameworks/react) and the [binding contract](/bindings).

## 3. Load and save files

The component ref (or the vanilla handle) exposes `load(bytes)` and `save()`. Loading detects DOCX and legacy `.doc`; password-protected files are not supported.

```ts
const bytes = new Uint8Array(await file.arrayBuffer());
await editor.load(bytes);

// After edits:
const blob = await editor.save(); // Blob of the serialized .docx
```

`dirty` and the `dirty-change` event report unsaved edits. See the [element API](/api) for every property, method and event.

## 4. Try it

- The [live demos](/demos) run the same editor in every framework adapter.
- To run the demo from a clone of the repository (Bun required):

```bash
bun install
bun run demo
```

## Next steps

- [Element API and events](/api)
- [Theming and light/dark sync](/theming)
- [Editing text](/editing) and [collaboration](/collaboration)
- [Architecture](/architecture) and [model units](/model-units)
