# Angular

El adaptador de Angular es un componente standalone (`spreadsheet-editor`) con inputs para las propiedades compartidas y salidas en camelCase.

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

La instancia del componente expone los métodos del [handle](/es/bindings#handle) (`load`, `save`, `download`, ...). Funciona con la detección de cambios sin zone.js.

Consulte el [contrato completo de los enlaces](/es/bindings) o [pruebe la demo de Angular](/demo-angular/).
