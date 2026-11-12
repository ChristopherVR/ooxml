# Bausteine

`docx-react-viewer` exportiert `useViewerBuildingBlocks`, `DocumentCanvas` sowie die optionalen Komponenten `Toolbar`, `PageNavigation` und `StatusBar`. Sie nutzen die vorhandene Engine des WordEditors für Bearbeitung, Layout, Verlauf und Serialisierung und folgen dem Aufbau des PowerPoint-Hooks `useViewerBuildingBlocks`: Der Hook liefert Props, und der Host rendert, was er möchte.

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, navigationProps, statusProps, loading, error } =
		useViewerBuildingBlocks({
			documentModel,
			canEdit,
			handle,
			onDirtyChange: (dirty) => console.log(dirty),
			// Host-owned dialogs: omit one and its action reports as not enabled.
			onOpenHyperlinkDialog: (request) => openMyLinkDialog(request),
			onOpenFindReplace: (request) => openMyFindBar(request),
			onOpenImagePicker: (request) => openMyFilePicker(request),
		});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{statusProps && <span>{statusProps.wordText}</span>}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

Rendern Sie die Zeichenfläche auch während des Ladens. Die Props sind null, bis sie eingehängt ist. Geben Sie dem Container der Zeichenfläche eine Höhe.

## Was der Hook zurückgibt {#what-the-hook-returns}

| Feld                       | Inhalt                                                                              |
| -------------------------- | ----------------------------------------------------------------------------------- |
| `canvasProps`              | Wird auf `DocumentCanvas` verteilt. Bindet den Editor ein und gibt ihn wieder frei. |
| `toolbarProps`             | Zustand von Auswahl und Formatierung sowie Bearbeitungs-Callbacks (unten).          |
| `navigationProps`          | Seite, Seitenanzahl, Ansichtsmodus und Callbacks zur Seitennavigation.              |
| `statusProps`              | Wortanzahl und Zoom mit Callbacks.                                                  |
| `loading`, `error`, `mode` | Ladezustand, der letzte Lade- oder Dokumentfehler, `'view'` oder `'edit'`.          |

### toolbarProps {#toolbarprops}

Zustand: `selection` (`{ from, to }`), `bold`, `italic`, `underline` (`null`, wenn die Auswahl gemischte Werte enthält), `fontFamily`, `fontSize` (`null`, wenn gemischt), `alignment` (`'left' | 'center' | 'right' | 'justify'` oder `null`), `list` (`'bullet' | 'number' | 'other' | null`), `canEdit`, `canUndo`, `canRedo` und übersetzte `labels`.

Callbacks, die jeweils `false` zurückgeben, wenn die Aktion nicht erlaubt ist (schreibgeschützt oder noch im Laden): `onBold`, `onItalic`, `onUnderline`, `onFontFamilyChange(family)`, `onFontSizeChange(points)`, `onAlignmentChange(alignment)`, `onListChange('bullet' | 'number' | 'remove')`, `onUndo`, `onRedo`. Der Zustand `list` und `onListChange` verwenden dieselben Werte `'bullet' | 'number'`.

### navigationProps {#navigationprops}

`page`, `pageCount`, `pageText` ("Page 2 of 5"), `pagesAvailable`, `canGoPrevious`, `canGoNext`, `viewMode` (`'draft' | 'print'`) sowie die Callbacks `onGoToPage(n)`, `onPreviousPage`, `onNextPage` und `onViewModeChange(mode)`. Seiten gibt es nur im Drucklayout: Im Entwurf ist `pagesAvailable` false, und die Seiten-Callbacks geben `false` zurück. Wechseln Sie mit `onViewModeChange('print')`.

### statusProps {#statusprops}

`words`, `wordText` ("42 words"), `zoom` (Prozent), `minZoom`, `maxZoom` sowie die Callbacks `onZoomChange(percent)`, `onZoomIn` und `onZoomOut`. Beim Scrollen wird auch `navigationProps` aktualisiert.

## Dialoge des Hosts {#host-dialogs}

Eine Oberfläche bindet keinen eigenen Dialog ein. Diese Abläufe gehören dem Host und sind wie die `onOpen...`-Callbacks von PowerPoint benannt. Übergeben Sie sie als Optionen. `toolbarProps` enthält dann einen passenden Öffner `open...` (`openHyperlinkDialog`, `openFindReplace`, `openImagePicker`), ein Flag `can...` und `hostWorkflows` (welche Callbacks übergeben wurden).

| Option                  | Anfrageobjekt                                                                    | Aktiviert, solange                        |
| ----------------------- | -------------------------------------------------------------------------------- | ----------------------------------------- |
| `onOpenHyperlinkDialog` | `{ href, text, apply(href, text?), remove() }` für den Link an der Auswahl       | bearbeitbar                               |
| `onOpenFindReplace`     | `{ search, findNext, findPrevious, replaceOne, replaceAll, status, canReplace }` | geladen (bei Schreibschutz bleibt Suchen) |
| `onOpenImagePicker`     | `{ insert(file) }`: Der Host wählt eine Datei aus und übergibt sie zurück        | bearbeitbar                               |

Wenn ein Callback fehlt, ist `canInsertLink`, `canFindReplace` oder `canInsertPicture` false, und der Öffner gibt `false` zurück. Anfragen prüfen die Berechtigung bei der Verwendung erneut. So kann ein Dialog, der bei einem Wechsel in den Schreibschutz offen geblieben ist, nicht mehr bearbeiten.

## Vorlesen {#read-aloud}

`toolbarProps` enthält eine Vorlesen-Steuerung, die die Implementierung des Editors nutzt (die
Sprachsynthese des Browsers; nichts wird gesendet). Sie liest die Auswahl oder, ohne Auswahl, von der
Einfügemarke bis zum Dokumentende, und funktioniert in einem schreibgeschützten Dokument: Sie ändert
nichts und markiert das Dokument nicht als geändert. Der Aufruf verschiebt weder Auswahl noch Fokus,
eine Schaltfläche Ihrer Oberfläche darf also zuvor den Fokus erhalten.

| Eigenschaft oder Callback | Bedeutung                                                                                                 |
| ------------------------- | --------------------------------------------------------------------------------------------------------- |
| `canReadAloud`            | Ein Dokument ist sichtbar und der Browser kann sprechen                                                   |
| `readingAloud`            | Dieser Editor liest gerade vor                                                                            |
| `readAloudIssue`          | `'unsupported'` (keine Sprachsynthese), `'empty'` (nichts zu lesen), `'error'` oder `null`                |
| `onReadAloud()`           | Startet das Vorlesen (bei Bedarf neu); gibt `'started'`, `'unsupported'`, `'empty'` oder `'error'` zurück |
| `onStopReadAloud()`       | Beendet die Wiedergabe dieses Editors; gibt zurück, ob sie lief                                           |

Der Browser hat nur eine Stimme. Wer in einem Editor startet, übernimmt sie von einem anderen,
dessen `readingAloud` dann `false` wird. Ein Editor hört auf, wenn sein Dokument bearbeitet oder
ersetzt wird, eine Datei geladen wird oder er entfernt wird. Formulieren Sie die Meldungen selbst;
die Bibliothek liefert nur die Codes. Die optionale `Toolbar` zeigt einen Schalter „Read aloud“ und
blendet ihn aus, wo keine Sprachausgabe möglich ist.

## Bibliothekskomponenten {#library-components}

`Toolbar` zeichnet Rückgängig und Wiederholen, Schriftart und Schriftgröße, Fett, Kursiv, Unterstrichen, vier Ausrichtungen, Aufzählungs- und nummerierte Listen sowie eine Schaltfläche für jeden Host-Ablauf, den Sie übergeben haben. `PageNavigation` zeichnet die vorherige Seite, "Page n of m", die nächste Seite und den Umschalter Entwurf bzw. Drucklayout. `StatusBar` zeichnet die Wortanzahl und eine Zoom-Auswahl. Alle sind optional und durch Props gesteuert. Jedes Steuerelement hat einen stabilen Knoten, sodass der Tastaturfokus Formatierungsaktualisierungen übersteht. Sie verwenden die gemeinsamen Office-Stile. Eigene Steuerelemente brauchen sie nie.

```tsx
{
	toolbarProps && <Toolbar {...toolbarProps} />;
}
{
	navigationProps && <PageNavigation {...navigationProps} />;
}
{
	statusProps && <StatusBar {...statusProps} />;
}
```

## Oberflächen im Schreibschutz {#read-only-surfaces}

Die Oberfläche ist **schreibgeschützt, bis der Host `canEdit: true` übergibt**. Das entspricht den Bausteinen von PowerPoint und Visio (der vollständige `WordEditor` ist dagegen standardmäßig bearbeitbar). Die Berechtigung wird angewendet, bevor die Zeichenfläche eingebunden wird. Deshalb beginnt `mode` mit `'view'`, und es gibt kein kurzes Aufblitzen eines bearbeitbaren Zustands. `readOnly` setzt sich gegenüber `canEdit: true` durch. Im Schreibschutz geben alle Bearbeitungs-Callbacks `false` zurück, und das Dokument kann sich nicht ändern. Die Leserin oder der Leser kann die Oberfläche trotzdem fokussieren, Text mit der Maus oder der Tastatur auswählen (Pfeiltasten, Umschalt mit Pfeiltasten, Pos1, Ende, Strg+A), kopieren und die Auswahl aus `toolbarProps.selection` lesen. Tastenkombinationen der Anwendung (Suchen, Speichern, Drucken) werden auf einer Oberfläche nicht installiert. Verbinden Sie sie mit Ihren eigenen Steuerelementen.

**Migration eines Hosts.** Bausteine waren früher standardmäßig bearbeitbar. Ein Host, der sich darauf verlassen hat, muss jetzt `canEdit: true` übergeben (im Optionsobjekt oder bei einem nativen Adapter in den reaktiven Optionen). Ohne diese Angabe öffnet sich die Oberfläche schreibgeschützt, und jeder Bearbeitungs-Callback gibt `false` zurück.

## Lebenszyklus {#lifecycle}

Verwenden Sie `bytes`, um eine Datei zu laden, oder die Handle-Methoden `load`, `save`, `download`, `markClean` und `dirty`. Ein Speichervorgang markiert ein Dokument erst dann als sauber, wenn der Host die Speicherung bestätigt. Die Zeichenfläche bindet keine Titelleiste, Multifunktionsleiste, Backstage, Statusleiste oder Anwendungsdialoge ein.

Jedes Framework-Paket exportiert außerdem `mountEditorSurface(host, options)` für native Lebenszyklus-Adapter. Dessen Binding stellt `getToolbarProps()`, `getNavigationProps()`, `getStatusProps()` und `subscribe(listener)` bereit, dazu das Editor-Handle, `update` und `destroy`. Die Optionen der Oberfläche schließen die Optionen aus, die nur für die Multifunktionsleiste gelten (`showToolbar`, `showThumbnails`, `hiddenActions`, `ribbonAddIns`, `onRibbonAddIn`). So kann ein TypeScript-Host keine Option übergeben, die ignoriert würde.

Die Composition-API benötigt die passende Laufzeitumgebung `ooxml-ui`, die zusammen mit diesen Ergänzungen veröffentlicht wird. Bestehende Importe des vollständigen Viewers bleiben mit der bisherigen veröffentlichten Laufzeitumgebung kompatibel. Eine ältere Laufzeitumgebung weist eine Composition-Anfrage ausdrücklich ab, statt versehentlich Anwendungsoberflächen einzubinden.

## Native Framework-Komposition {#native-framework-composition}

Framework-Adapter liegen in ihrem eigenen Framework-Paket. Sie teilen sich den Helfer für die Lebensdauer der Engine und die Steuerelement-Beschreibungen (`bindings/src/controls.ts`), die die optionalen Komponenten zeichnen. Jeder Adapter stellt `toolbarProps`, `navigationProps`, `statusProps`, `loading`, `error`, `mode` und `canvasProps` bereit, dazu das Editor-`handle` (React übernimmt ein `handle`-Ref als Option statt es zurückzugeben) und die Komponenten `Toolbar`, `PageNavigation` und `StatusBar` (Angular: `ToolbarComponent`, `PageNavigationComponent`, `StatusBarComponent`; Svelte: `docx-svelte-viewer/Toolbar`, `/PageNavigation`, `/StatusBar`). Ihr Zustand nutzt die üblichen reaktiven Primitive des Frameworks, und ihre Zeichenflächen-Komponenten binden die vorhandene Word-Dokumentoberfläche ein.

### Vue {#vue}

```vue
<script setup lang="ts">
import { computed } from 'vue';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-vue-viewer';
const props = defineProps<{ documentModel: import('docx-core').DocumentModel; canEdit: boolean }>();
const { canvasProps, toolbarProps, error, handle } = useViewerBuildingBlocks(
	computed(() => ({ documentModel: props.documentModel, canEdit: props.canEdit })),
);
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<DocumentCanvas v-bind="canvasProps" class="document-surface" />
	<p v-if="error" role="alert">{{ error.message }}</p>
</template>
```

Die Composable beobachtet einen Ref, einen berechneten Ref oder einen Getter und gibt die Engine mit ihrem Scope frei. `toolbarProps`, `loading`, `error` und `mode` sind schreibgeschützte berechnete Refs. `handle` stellt die vorhandene imperative Editor-API bereit.

### Angular {#angular}

Stellen Sie pro Host-Komponente einen Dienst `ViewerBuildingBlocks` bereit. Signale, die im Template gelesen werden, nehmen an der Änderungserkennung von Angular teil, auch in zonenlosen Anwendungen. `DestroyRef` gibt die Engine frei, sobald ihr Provider zerstört wird.

```ts
import { Component, signal } from '@angular/core';
import { createDocument } from 'docx-core';
import {
	createViewerBuildingBlocks,
	ViewerBuildingBlocks,
	DocumentCanvasComponent,
	ToolbarComponent,
} from 'docx-angular-viewer';

@Component({
	selector: 'embedded-word',
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent],
	providers: [ViewerBuildingBlocks],
	template: `
		@if (blocks.toolbarProps(); as toolbar) {
			<word-toolbar [toolbarProps]="toolbar" />
		}
		<word-document-canvas [canvasProps]="blocks.canvasProps" />
	`,
})
export class EmbeddedWord {
	readonly canEdit = signal(true);
	readonly model = createDocument();
	readonly blocks = createViewerBuildingBlocks(() => ({
		documentModel: this.model,
		canEdit: this.canEdit(),
	}));
}
```

Der Dienst stellt schreibgeschützte Signale `toolbarProps`, `loading`, `error` und `mode`, ein stabiles `handle` sowie typisierte Event-Emitter für `documentChange`, `documentError`, `dirtyChange` und `pageChange` bereit. Abonnieren Sie diese Emitter oder übergeben Sie die vorhandenen Optionen als Callbacks. Für imperative Hosts injizieren Sie den Dienst und rufen `update(options)` auf, statt den Helfer zu verwenden.

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
  import { writable } from 'svelte/store';
  import { createDocument } from 'docx-core';
  import { createViewerBuildingBlocks } from 'docx-svelte-viewer/runtime';
  import DocumentCanvas from 'docx-svelte-viewer/DocumentCanvas';
  import Toolbar from 'docx-svelte-viewer/Toolbar';
  const options = writable({ documentModel: createDocument(), canEdit: true });
  const blocks = createViewerBuildingBlocks(options);
  const { toolbarProps, error } = blocks;
</script>
{#if $toolbarProps}<Toolbar toolbarProps={$toolbarProps} />{/if}
<DocumentCanvas {...blocks.canvasProps} />
{#if $error}<p role="alert">{$error.message}</p>{/if}
```

Rufen Sie die Factory während der Initialisierung der Komponente auf. Sie akzeptiert einen lesbaren Optionsspeicher oder ein anfängliches Optionsobjekt. `blocks.update(next)` unterstützt außerdem Aktualisierungen über `$effect` mit Runes. Snapshots sind lesbare Speicher. `documentCanvas` wird als Action für eigenes Host-Markup exportiert. Das Zerstören der Komponente beendet das Abonnement der Optionen und zerstört die Engine.

### Solid {#solid}

```tsx
import { createSignal, Show } from 'solid-js';
import { createDocument } from 'docx-core';
import { createViewerBuildingBlocks, DocumentCanvas, Toolbar } from 'docx-solid-viewer';
function EmbeddedWord() {
	const [canEdit, setCanEdit] = createSignal(true);
	const model = createDocument();
	const blocks = createViewerBuildingBlocks(() => ({ documentModel: model, canEdit: canEdit() }));
	return (
		<>
			<Show when={blocks.toolbarProps()}>{(toolbar) => <Toolbar toolbarProps={toolbar()} />}</Show>
			<DocumentCanvas {...blocks.canvasProps} />
		</>
	);
}
```

Die Factory beobachtet ihren Optionszugriff in einem Effect und gibt Signalzugriffe für `toolbarProps`, `loading`, `error` und `mode` zurück. Der Besitzer steuert das Aufräumen über `onCleanup`. Optionale Steuerelemente der Multifunktionsleiste behalten ihre DOM-Knoten, während sich der Formatierungszustand ändert, sodass der Tastaturfokus erhalten bleibt. Alle Adapter nehmen `bytes` reaktiv entgegen, verwerfen fehlgeschlagene Ladevorgänge, wenn ein neues Modell oder Laden eintrifft, und behalten ihr Handle stabil, wenn die Zeichenfläche ersetzt wird.

### Vanilla {#vanilla}

`docx-vanilla-viewer` exportiert die gemeinsame Sitzung als `createViewerBuildingBlocks(options)` (`canvasProps`, `handle`, `state`, `subscribe`, `update`, `destroy`) sowie Mount-Helfer, die einen Disposer zurückgeben: `mountDocumentCanvas(host, blocks)`, `mountToolbar(host, blocks)`, `mountPageNavigation(host, blocks)` und `mountStatusBar(host, blocks)`.

```ts
import {
	createViewerBuildingBlocks,
	mountDocumentCanvas,
	mountToolbar,
	mountStatusBar,
} from 'docx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ documentModel, canEdit: true });
const stops = [
	mountDocumentCanvas(document.querySelector('#page')!, blocks),
	mountToolbar(document.querySelector('#tools')!, blocks),
	mountStatusBar(document.querySelector('#status')!, blocks),
];
blocks.subscribe((state) => console.log(state.toolbarProps?.bold)); // your own controls
// later
stops.forEach((stop) => stop());
blocks.destroy();
```

## Gemeinsames Styling {#shared-styling}

Die optionalen nativen Werkzeugleisten-Komponenten verwenden die gemeinsamen Styles von `ooxml-ui`, einschließlich Abstände, Hintergründe, Rahmen, Zustände für gedrückt und deaktiviert sowie Fokusringe für die Tastatur. Platzieren Sie Werkzeugleiste und Zeichenfläche unter einem gemeinsamen Vorfahren mit `data-office-theme` und passen Sie die Variablen `--office-*` dort an. Änderungen wirken sofort in allen Frameworks. Vorhandene Überschreibungen von `--dve-*` beim Vorfahren werden weiterhin unterstützt. Vom Host erzeugte Steuerelemente behalten ihr eigenes Styling, es sei denn, der Host verwendet den gemeinsamen Helfer `prepareOfficeUiBlock(element, 'toolbar', 'docx')` aus `ooxml-ui/theme`.
