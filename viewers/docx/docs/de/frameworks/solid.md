# SolidJS

Der SolidJS-Adapter bindet das gemeinsame Element `<docx-editor>` ein und leitet Modelländerungen über `onDocumentChange` weiter.

```tsx
import { WordEditor } from 'docx-solid-viewer';

function Editor(props) {
	return (
		<WordEditor
			documentModel={props.model}
			onDocumentChange={props.setModel}
			editorRef={(handle) => props.setHandle(handle)}
		/>
	);
}
```

Der Adapter ruft `editorRef` mit einem Handle auf, das `load(bytes)` und `save()` bereitstellt. Siehe den [vollständigen Binding-Vertrag](/de/bindings) oder [probieren Sie die Solid-Demo aus](/demo-solid/).
