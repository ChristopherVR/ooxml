# SolidJS

L'adaptateur SolidJS monte l'élément partagé `<docx-editor>` et transmet les changements de modèle par `onDocumentChange`.

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

L'adaptateur appelle `editorRef` avec un handle exposant `load(bytes)` et `save()`. Consultez le [contrat complet des adaptateurs](/fr/bindings) ou [essayez la démo Solid](/demo-solid/).
