# Angular

L’adaptateur Angular est un composant autonome (`spreadsheet-editor`) doté d’entrées pour les props communes et de sorties en camelCase.

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

L’instance du composant expose les méthodes de la [poignée](/fr/bindings#handle) (`load`, `save`, `download`, ...). Il fonctionne avec la détection de changements sans zone.

Consultez le [contrat complet des liaisons](/fr/bindings) ou [essayez la démo Angular](/demo-angular/).
