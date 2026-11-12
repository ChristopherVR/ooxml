# JavaScript nativo

Importe el adaptador de montaje vanilla y proporcione un elemento anfitrión. El manejador devuelto incluye load, save, update, destroy y el elemento personalizado.

```ts
import { mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

El elemento personalizado también puede registrarse directamente con `registerDocxEditor` desde `docx-vanilla-viewer`. Consulte el [contrato completo de los adaptadores](/es/bindings) o [pruebe la demo de JavaScript nativo](/demo-vanilla/).
