# Primeros pasos

## Ejecutar la demo desde el código fuente

Necesita [Bun](https://bun.sh/) 1.4 y Node.js 24.

```bash
git clone https://github.com/ChristopherVR/ooxml.git
bun install   # once, at the repository root (one workspace, one lockfile)
cd viewers/xlsx
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

La página de inicio abre un archivo `.xlsx`, `.xlsm`, `.xls` heredado o `.csv` (suéltelo o búsquelo), crea un libro nuevo o abre el libro de ejemplo. Los archivos se procesan únicamente en el navegador.

## Elegir un paquete

Instale un paquete de editor autónomo para su framework. Cada uno incluye el componente web `<xlsx-editor>` y su adaptador, depende del modelo de libro (`@christophervr/xlsx-core`), de `ooxml-core` y de `ooxml-ui`, y reexporta la API del modelo (`createWorkbook`, `loadWorkbook`, `saveWorkbook`, ...), de modo que a una aplicación le basta con una sola instalación y una sola ruta de importación.

| Framework | Paquete                            | Componente                     |
| --------- | ---------------------------------- | ------------------------------ |
| React     | `@christophervr/xlsx-react-viewer` | `SpreadsheetEditor`            |
| Vue 3     | `xlsx-vue-viewer`                  | `SpreadsheetEditor`            |
| Angular   | `xlsx-angular-viewer`              | `SpreadsheetEditorComponent`   |
| Svelte 5  | `xlsx-svelte-viewer`               | default export, `/runtime`     |
| Solid     | `xlsx-solid-viewer`                | `SpreadsheetEditor`            |
| Vanilla   | `xlsx-vanilla-viewer`              | `mountEditor`, `<xlsx-editor>` |

Instale `@christophervr/xlsx-core` por separado solo para el uso sin interfaz (analizar, calcular y guardar sin un editor).

## Montar un editor

```tsx
import { useState } from 'react';
import { createWorkbook, SpreadsheetEditor } from '@christophervr/xlsx-react-viewer';

export function Editor() {
	const [workbook, setWorkbook] = useState(() => createWorkbook());
	return (
		<div style={{ height: '100vh' }}>
			<SpreadsheetEditor workbook={workbook} onWorkbookChange={setWorkbook} />
		</div>
	);
}
```

Dé una altura al contenedor: la cuadrícula lo llena y se desplaza dentro de él. Para abrir un archivo, pase sus bytes (`bytes={uint8array}` con `fileName`), una URL (`src="/files/budget.xlsx"`) o llame a `load(bytes, fileName)` en el manejador del componente.

## Guardar

`save()` se resuelve con un `Blob` `.xlsx`, `saveBytes('csv')` exporta la hoja activa como CSV y `download(name?)` guarda e inicia una descarga del navegador. `dirty` indica si hay cambios sin guardar; llame a `markClean()` después de persistir usted mismo un Blob guardado. Un archivo `.xls` heredado se guarda como `.xlsx`.

Siguiente: los [enlaces de framework](/bindings), la [API del elemento](/es/api), los [temas](/theming) y las [funciones y limitaciones](/es/features).
