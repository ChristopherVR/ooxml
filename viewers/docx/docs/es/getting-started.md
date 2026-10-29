# Primeros pasos

docx-viewer es un único componente web `<docx-editor>` más seis adaptadores de framework muy ligeros. Esta página instala un adaptador, monta el editor y carga y guarda un documento. Es estable según la [política de estabilidad](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md), con una [garantía de conservación](/preservation), pero no equivale a Microsoft Word (consulte la [hoja de ruta de compatibilidad](/parity-roadmap)).

## 1. Instalación

Instale el paquete autónomo de su framework. Incluye la interfaz del editor y el adaptador, y trae consigo `docx-core` (el modelo de documento).

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

Instale `docx-core` por separado solo para el uso sin interfaz (analizar y serializar sin editor).

## 2. Montar el editor

Todos los adaptadores montan el mismo elemento `<docx-editor>`. Pase un `DocumentModel`, escuche los cambios y asigne una altura al contenedor.

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

Los demás frameworks usan las mismas opciones: consulte las [guías por framework](/frameworks/react) y el [contrato de los adaptadores](/bindings).

## 3. Cargar y guardar archivos

La referencia al componente (o el manejador de la versión vanilla) expone `load(bytes)` y `save()`. La carga detecta DOCX y el formato `.doc` heredado; los archivos protegidos con contraseña no son compatibles.

```ts
const bytes = new Uint8Array(await file.arrayBuffer());
await editor.load(bytes);

// After edits:
const blob = await editor.save(); // Blob of the serialized .docx
```

`dirty` y el evento `dirty-change` informan de las ediciones sin guardar. Consulte la [API del elemento](/es/api) para ver todas las propiedades, métodos y eventos.

## 4. Pruébelo

- Las [demos en vivo](/demos) ejecutan el mismo editor en todos los adaptadores de framework.
- Para ejecutar la demo desde un clon del repositorio (requiere Bun):

```bash
bun install
bun run demo
```

## Siguientes pasos

- [API del elemento y eventos](/es/api)
- [Temas y sincronización claro/oscuro](/theming)
- [Edición de texto](/es/editing) y [colaboración](/es/collaboration)
- [Arquitectura](/es/architecture) y [unidades del modelo](/model-units)
