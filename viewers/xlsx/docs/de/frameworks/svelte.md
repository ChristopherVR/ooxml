# Svelte

Verwenden Sie den Svelte-5-Adapter als Komponente und behandeln Sie Änderungen an der Arbeitsmappe über die Callback-Props.

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

Das Wurzelelement des Pakets ist die Komponente (Default-Export). Die reinen JavaScript-Hilfsfunktionen (`mountEditor`, `loadWorkbook`, `createWorkbook` und die gemeinsamen Typen) liegen unter `xlsx-svelte-viewer/runtime`. Exporte der Komponente: `load`, `newWorkbook`, `save`, `saveBytes`, `download`, `markClean`, `select`, `getSelection`, `setActiveSheet`, die aktuellen Eigenschaften `element` und `dirty` (`isDirty()` und `getElement()` bleiben als veraltete Aliase erhalten).

Siehe die [vollständige Binding-Vertragsbeschreibung](/de/bindings) oder [probieren Sie die Svelte-Demo aus](/demo-svelte/).
