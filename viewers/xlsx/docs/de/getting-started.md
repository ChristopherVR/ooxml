# Erste Schritte

## Die Demo aus dem Quellcode starten

Sie benötigen [Bun](https://bun.sh/) 1.3 und Node.js 24.

```bash
git clone https://github.com/ChristopherVR/ooxml.git
bun install   # once, at the repository root (one workspace, one lockfile)
cd viewers/xlsx
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

Die Startseite öffnet eine `.xlsx`-, `.xlsm`-, ältere `.xls`- oder `.csv`-Datei (per Drag-and-drop oder über die Dateiauswahl), startet eine neue Arbeitsmappe oder öffnet die Beispielarbeitsmappe. Dateien werden ausschließlich im Browser verarbeitet.

## Ein Paket auswählen

Installieren Sie ein eigenständiges Editor-Paket für Ihr Framework. Jedes Paket enthält die Web Component `<xlsx-editor>` samt Adapter, hängt vom Arbeitsmappenmodell (`@christophervr/xlsx-core`), von `ooxml-core` und von `ooxml-ui` ab und exportiert die Modell-API (`createWorkbook`, `loadWorkbook`, `saveWorkbook`, ...) erneut. Eine einzige Installation und ein einziger Importpfad genügen daher für eine Anwendung.

| Framework | Paket                              | Komponente                     |
| --------- | ---------------------------------- | ------------------------------ |
| React     | `@christophervr/xlsx-react-viewer` | `SpreadsheetEditor`            |
| Vue 3     | `xlsx-vue-viewer`                  | `SpreadsheetEditor`            |
| Angular   | `xlsx-angular-viewer`              | `SpreadsheetEditorComponent`   |
| Svelte 5  | `xlsx-svelte-viewer`               | default export, `/runtime`     |
| Solid     | `xlsx-solid-viewer`                | `SpreadsheetEditor`            |
| Vanilla   | `xlsx-vanilla-viewer`              | `mountEditor`, `<xlsx-editor>` |

Installieren Sie `@christophervr/xlsx-core` nur dann einzeln, wenn Sie ohne Editor arbeiten (Mappen parsen, berechnen und speichern).

## Einen Editor einbinden

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

Geben Sie dem Container eine Höhe: Das Raster füllt ihn aus und scrollt darin. Um eine Datei zu öffnen, übergeben Sie ihre Bytes (`bytes={uint8array}` mit `fileName`), eine URL (`src="/files/budget.xlsx"`) oder rufen Sie `load(bytes, fileName)` am Komponenten-Handle auf.

## Speichern

`save()` liefert einen `.xlsx`-`Blob`, `saveBytes('csv')` exportiert das aktive Blatt als CSV, und `download(name?)` speichert und startet einen Browser-Download. `dirty` meldet ungespeicherte Änderungen; rufen Sie `markClean()` auf, nachdem Sie einen gespeicherten Blob selbst abgelegt haben. Eine ältere `.xls`-Datei wird als `.xlsx` gespeichert.

Weiter geht es mit den [Framework-Bindings](/bindings), der [Element-API](/de/api), den [Themes](/theming) und den [Funktionen und Einschränkungen](/de/features).
