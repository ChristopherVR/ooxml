# Solid

`visio-solid-viewer` stellt den gemeinsamen `<visio-viewer>` als native Solid-Integration bereit.

```bash
npm install visio-solid-viewer solid-js
```

```tsx
import { VisioViewer } from 'visio-solid-viewer';
import type { VisioDocument } from 'visio-solid-viewer';

export function Diagram(props: { diagram: VisioDocument }) {
	return (
		<VisioViewer
			document={props.diagram}
			style={{ height: '600px' }}
			events={{ 'document-error': console.error }}
		/>
	);
}
```

Übergeben Sie `viewerRef`, um das Handle beim Mounten zu erhalten (und `undefined` beim Aushängen).

Die Eigenschaften (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) und das Handle sind in jedem Framework identisch: siehe den [Binding-Vertrag](/de/bindings) und die [Viewer-API](/de/api). Um stattdessen eine eigene Oberfläche aus den Teilen des Viewers zusammenzusetzen, siehe die [Bausteine](/de/building-blocks); um eine Zeichnung zu teilen, siehe [Zusammenarbeit](/de/collaboration). [Die Solid-Demo ausprobieren](/demo-solid/){target="_self"}.
