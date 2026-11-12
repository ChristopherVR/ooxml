# React

Der React-Adapter bindet das gemeinsame Element `<docx-editor>` ein und leitet Modellaktualisierungen über `onDocumentChange` weiter.

```tsx
import { WordEditor } from 'docx-react-viewer';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

Verwenden Sie die Komponenten-Ref für `load(bytes)` und `save()`. Siehe den [vollständigen Binding-Vertrag](/de/bindings) oder [probieren Sie die React-Demo aus](/demo/).
