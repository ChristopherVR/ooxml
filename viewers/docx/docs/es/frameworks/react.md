# React

El adaptador de React monta el elemento compartido `<docx-editor>` y reenvía las actualizaciones del modelo a través de `onDocumentChange`.

```tsx
import { WordEditor } from 'docx-react-viewer';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

Use la referencia del componente para `load(bytes)` y `save()`. Consulte el [contrato completo de los adaptadores](/es/bindings) o [pruebe la demo de React](/demo/).
