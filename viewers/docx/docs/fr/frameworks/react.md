# React

L'adaptateur React monte l'élément partagé `<docx-editor>` et transmet les mises à jour du modèle par `onDocumentChange`.

```tsx
import { WordEditor } from 'docx-react-viewer';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

Utilisez la référence du composant pour `load(bytes)` et `save()`. Consultez le [contrat complet des adaptateurs](/fr/bindings) ou [essayez la démo React](/demo/).
