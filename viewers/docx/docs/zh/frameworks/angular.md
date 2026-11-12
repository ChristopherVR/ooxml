# Angular

导入独立组件，并绑定其文档输入和变更输出。

```ts
import { WordEditorComponent } from 'docx-angular-viewer';

@Component({
	standalone: true,
	imports: [WordEditorComponent],
	template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
	model = initialDocument;
}
```

参见[完整的绑定契约](/zh/bindings)或[试用 Angular 演示](/demo-angular/)。
