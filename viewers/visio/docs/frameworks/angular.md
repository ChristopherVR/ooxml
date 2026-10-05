# Angular

`visio-angular-viewer` gives you the shared `<visio-viewer>` as a native Angular integration.

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

Outputs are `documentLoad`, `documentError`, `pageChange`, `zoomChange` and `shapeSelect`; `documentChange` reports edits. The complete `events` callback map is also accepted as an input.

The properties (`document`, `pageIndex`, `zoom`, `showToolbar`, `events`) and the handle are identical in every framework: see the [binding contract](/bindings) and the [viewer API](/api). [Try the Angular demo](/demo-angular/){target="_self"}.
