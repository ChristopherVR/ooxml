# React

`visio-react-viewer` stellt den gemeinsamen `<visio-viewer>` als native React-Integration bereit.

```bash
npm install visio-react-viewer react
```

```tsx
import { useRef } from 'react';
import { VisioViewer, type ViewerHandle, type VisioDocument } from 'visio-react-viewer';

export function Diagram({ diagram }: { diagram: VisioDocument }) {
	const ref = useRef<ViewerHandle>(null);
	return (
		<VisioViewer
			ref={ref}
			document={diagram}
			style={{ height: 600 }}
			events={{ 'document-error': console.error }}
		/>
	);
}
```

Eine weitergeleitete `ref` stellt das `ViewerHandle` bereit. Die React-Props `className` und `style` gestalten den Host, der eine Höhe benötigt.

Die Eigenschaften (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) und das Handle sind in jedem Framework identisch: siehe den [Binding-Vertrag](/de/bindings) und die [Viewer-API](/de/api). Um stattdessen eine eigene Oberfläche aus den Teilen des Viewers zusammenzusetzen, siehe die [Bausteine](/de/building-blocks); um eine Zeichnung zu teilen, siehe [Zusammenarbeit](/de/collaboration). [Die React-Demo ausprobieren](/demo-react/){target="_self"}.
