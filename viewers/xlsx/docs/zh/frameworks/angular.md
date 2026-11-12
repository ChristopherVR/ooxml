# Angular

Angular 适配器是一个独立组件（`spreadsheet-editor`），通过输入属性接收共享属性，并使用驼峰命名的输出。

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

组件实例公开了[句柄](/zh/bindings#handle)的方法（`load`、`save`、`download` 等）。它可在无 Zone 的变更检测下工作。

参见[完整的绑定契约](/zh/bindings)，或[试用 Angular 演示](/demo-angular/)。
