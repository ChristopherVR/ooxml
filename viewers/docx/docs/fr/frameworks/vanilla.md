# JavaScript natif

Importez l'adaptateur de montage natif et fournissez un élément hôte. Le handle renvoyé comprend load, save, update, destroy et l'élément personnalisé.

```ts
import { mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

L'élément personnalisé peut aussi être enregistré directement avec `registerDocxEditor` depuis `docx-vanilla-viewer`. Consultez le [contrat complet des adaptateurs](/fr/bindings) ou [essayez la démo JavaScript natif](/demo-vanilla/).
