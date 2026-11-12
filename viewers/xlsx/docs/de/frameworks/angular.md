# Angular

Der Angular-Adapter ist eine eigenständige Komponente (`spreadsheet-editor`) mit Eingaben für die gemeinsamen Props und Ausgaben in camelCase.

```ts
import { Component } from '@angular/core';
import { createWorkbook, SpreadsheetEditorComponent } from 'xlsx-angular-viewer';

@Component({
	selector: 'app-editor',
	standalone: true,
	imports: [SpreadsheetEditorComponent],
	template: `<spreadsheet-editor
		#editor
		[workbook]="workbook"
		[readOnly]="false"
		(workbookChange)="workbook = $event"
		(dirtyChange)="dirty = $event"
	/>`,
})
export class EditorComponent {
	workbook = createWorkbook();
	dirty = false;
}
```

Die Komponenteninstanz stellt die Methoden des [Handles](/de/bindings#handle) bereit (`load`, `save`, `download` und weitere). Sie funktioniert mit Change Detection ohne Zone.js.

Siehe die [vollständige Binding-Vertragsbeschreibung](/de/bindings) oder [probieren Sie die Angular-Demo aus](/demo-angular/).
