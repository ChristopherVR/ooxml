# JavaScript vanilla

`visio-vanilla-viewer` le ofrece el `<visio-viewer>` compartido como integración nativa de JavaScript vanilla.

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

`mountViewer` devuelve el enlace compartido completo con `update()` y un `destroy()` idempotente.

Las propiedades (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) y el handle son idénticos en todos los frameworks: consulte el [contrato de enlaces](/es/bindings) y la [API del visor](/es/api). Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/es/building-blocks); para compartir un dibujo, consulte [colaboración](/es/collaboration). [Pruebe la demostración de JavaScript vanilla](/demo/){target="_self"}.
