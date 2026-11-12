# Angular

Importez le composant autonome et liez son entrée de document et sa sortie de changement.

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

Consultez le [contrat complet des adaptateurs](/fr/bindings) ou [essayez la démo Angular](/demo-angular/).
