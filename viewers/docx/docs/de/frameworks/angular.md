# Angular

Importieren Sie die eigenständige Komponente und binden Sie ihren Dokument-Input und ihren Änderungs-Output.

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

Siehe den [vollständigen Binding-Vertrag](/de/bindings) oder [probieren Sie die Angular-Demo aus](/demo-angular/).
