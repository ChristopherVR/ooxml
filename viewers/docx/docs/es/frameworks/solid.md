# SolidJS

El adaptador de SolidJS monta el elemento compartido `<docx-editor>` y reenvía los cambios del modelo a través de `onDocumentChange`.

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

El adaptador llama a `editorRef` con un manejador que expone `load(bytes)` y `save()`. Consulte el [contrato completo de los adaptadores](/es/bindings) o [pruebe la demo de Solid](/demo-solid/).
