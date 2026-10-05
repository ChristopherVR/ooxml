# Angular

The Angular adapter is a standalone component (`spreadsheet-editor`) with inputs for the shared props and camelCase outputs.

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

The component instance exposes the [handle](/bindings#handle) methods (`load`, `save`, `download`, ...). It works with zoneless change detection.

See the [complete binding contract](/bindings) or [try the Angular demo](/demo-angular/).
