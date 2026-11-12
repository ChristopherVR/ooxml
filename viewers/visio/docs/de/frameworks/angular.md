# Angular

`visio-angular-viewer` stellt den gemeinsamen `<visio-viewer>` als native Angular-Integration bereit.

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

Die Outputs sind `documentLoad`, `documentChange`, `documentError`, `pageChange`, `zoomChange`, `shapeSelect`, `selectionChange`, `ribbonAddIn`, `collaborationChange` und `fileCommand`. `fileCommand` meldet nur einen Datei-Befehl: Um ihn abzubrechen, geben Sie `false` aus dem Callback `events['file-command']` zurück. Die vollständige Callback-Map `events` wird ebenfalls als Input akzeptiert.

Die Eigenschaften (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) und das Handle sind in jedem Framework identisch: siehe den [Binding-Vertrag](/de/bindings) und die [Viewer-API](/de/api). Um stattdessen eine eigene Oberfläche aus den Teilen des Viewers zusammenzusetzen, siehe die [Bausteine](/de/building-blocks); um eine Zeichnung zu teilen, siehe [Zusammenarbeit](/de/collaboration). [Die Angular-Demo ausprobieren](/demo-angular/){target="_self"}.
