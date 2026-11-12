# Angular

Importe el componente standalone y enlace su entrada de documento y su salida de cambios.

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

Consulte el [contrato completo de los adaptadores](/es/bindings) o [pruebe la demo de Angular](/demo-angular/).
