# Angular

`visio-angular-viewer` le ofrece el `<visio-viewer>` compartido como integración nativa de Angular.

```bash
npm install visio-angular-viewer @angular/core
```

```ts
import { Component, Input } from '@angular/core';
import { VisioViewerComponent } from 'visio-angular-viewer';
import type { VisioDocument } from 'visio-angular-viewer';

@Component({
	selector: 'app-diagram',
	imports: [VisioViewerComponent],
	template: '<visio-viewer-host [document]="diagram" (documentError)="onError($event)" />',
	styles: ['visio-viewer-host { display: block; height: 600px }'],
})
export class Diagram {
	@Input() diagram: VisioDocument | null = null;
	onError(error: Error) {
		console.error(error);
	}
}
```

Las salidas son `documentLoad`, `documentChange`, `documentError`, `pageChange`, `zoomChange`, `shapeSelect`, `selectionChange`, `ribbonAddIn`, `collaborationChange` y `fileCommand`. `fileCommand` solo informa de un comando de Archivo: para cancelarlo, devuelva `false` desde el callback `events['file-command']`. El mapa completo de callbacks `events` también se acepta como entrada.

Las propiedades (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) y el handle son idénticos en todos los frameworks: consulte el [contrato de enlaces](/es/bindings) y la [API del visor](/es/api). Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/es/building-blocks); para compartir un dibujo, consulte [colaboración](/es/collaboration). [Pruebe la demostración de Angular](/demo-angular/){target="_self"}.
