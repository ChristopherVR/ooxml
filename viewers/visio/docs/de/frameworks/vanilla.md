# Vanilla JavaScript

`visio-vanilla-viewer` stellt den gemeinsamen `<visio-viewer>` als native Integration für Vanilla JavaScript bereit.

```bash
npm install visio-vanilla-viewer
```

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	document: diagram,
	zoom: 1,
	events: { 'document-error': console.error },
});

viewer.update({ pageIndex: 1 });
await viewer.load(file);
viewer.destroy(); // idempotent
```

`mountViewer` gibt das vollständige gemeinsame Binding mit `update()` und einem idempotenten `destroy()` zurück.

Die Eigenschaften (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) und das Handle sind in jedem Framework identisch: siehe den [Binding-Vertrag](/de/bindings) und die [Viewer-API](/de/api). Um stattdessen eine eigene Oberfläche aus den Teilen des Viewers zusammenzusetzen, siehe die [Bausteine](/de/building-blocks); um eine Zeichnung zu teilen, siehe [Zusammenarbeit](/de/collaboration). [Die Demo für Vanilla JavaScript ausprobieren](/demo/){target="_self"}.
