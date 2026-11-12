# React

`visio-react-viewer` le ofrece el `<visio-viewer>` compartido como integración nativa de React.

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

Una `ref` reenviada expone el `ViewerHandle`. Las props `className` y `style` de React dan estilo al host, que necesita una altura.

Las propiedades (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) y el handle son idénticos en todos los frameworks: consulte el [contrato de enlaces](/es/bindings) y la [API del visor](/es/api). Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/es/building-blocks); para compartir un dibujo, consulte [colaboración](/es/collaboration). [Pruebe la demostración de React](/demo-react/){target="_self"}.
