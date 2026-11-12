# Solid

`visio-solid-viewer` le ofrece el `<visio-viewer>` compartido como integración nativa de Solid.

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

Pase `viewerRef` para recibir el handle al montarse (y `undefined` al desmontarse).

Las propiedades (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) y el handle son idénticos en todos los frameworks: consulte el [contrato de enlaces](/es/bindings) y la [API del visor](/es/api). Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/es/building-blocks); para compartir un dibujo, consulte [colaboración](/es/collaboration). [Pruebe la demostración de Solid](/demo-solid/){target="_self"}.
