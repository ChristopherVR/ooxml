# Vanilla JavaScript

Importieren Sie den Vanilla-Mount-Adapter und stellen Sie ein Host-Element bereit. Das zurückgegebene Handle umfasst load, save, update, destroy und das Custom Element.

```ts
import { mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

Das Custom Element kann auch direkt mit `registerDocxEditor` aus `docx-vanilla-viewer` registriert werden. Siehe den [vollständigen Binding-Vertrag](/de/bindings) oder [probieren Sie die Vanilla-Demo aus](/demo-vanilla/).
