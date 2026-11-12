# Svelte

Utilisez l’adaptateur Svelte 5 comme un composant et traitez les changements de classeur avec les props de rappel.

```svelte
<script>
	import XlsxEditor from 'xlsx-svelte-viewer';
	import { createWorkbook } from 'xlsx-svelte-viewer/runtime';

	let workbook = $state.raw(createWorkbook());
	let editor;
</script>

<button onclick={() => editor.download('Budget.xlsx')}>Download</button>
<XlsxEditor bind:this={editor} {workbook} onworkbookchange={(next) => (workbook = next)} />
```

La racine du paquet est le composant (export par défaut). Les utilitaires en JavaScript simple (`mountEditor`, `loadWorkbook`, `createWorkbook` et les types communs) se trouvent dans `xlsx-svelte-viewer/runtime`. Exports du composant : `load`, `newWorkbook`, `save`, `saveBytes`, `download`, `markClean`, `select`, `getSelection`, `setActiveSheet`, ainsi que les propriétés actives `element` et `dirty` (`isDirty()` et `getElement()` restent disponibles comme alias dépréciés).

Consultez le [contrat complet des liaisons](/fr/bindings) ou [essayez la démo Svelte](/demo-svelte/).
