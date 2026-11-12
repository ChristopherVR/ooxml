# Svelte

Use el adaptador de Svelte 5 como componente y gestione los cambios del libro con las props de callback.

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

La raíz del paquete es el componente (exportación por defecto). Los auxiliares en JavaScript plano (`mountEditor`, `loadWorkbook`, `createWorkbook` y los tipos compartidos) están en `xlsx-svelte-viewer/runtime`. Exports del componente: `load`, `newWorkbook`, `save`, `saveBytes`, `download`, `markClean`, `select`, `getSelection`, `setActiveSheet`, y las propiedades `element` y `dirty` en vivo (`isDirty()` y `getElement()` siguen disponibles como alias obsoletos).

Consulte el [contrato completo de los enlaces](/es/bindings) o [pruebe la demo de Svelte](/demo-svelte/).
