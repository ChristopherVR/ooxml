# Svelte

将 Svelte 5 适配器作为组件使用，并通过回调属性处理工作簿变化。

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

软件包根部导出的是组件（默认导出）。纯 JavaScript 辅助函数（`mountEditor`、`loadWorkbook`、`createWorkbook` 以及共享类型）位于 `xlsx-svelte-viewer/runtime` 下。组件导出包括：`load`、`newWorkbook`、`save`、`saveBytes`、`download`、`markClean`、`select`、`getSelection`、`setActiveSheet`，以及实时的 `element` 和 `dirty` 属性（`isDirty()` 和 `getElement()` 作为已弃用的别名保留）。

参见[完整的绑定契约](/zh/bindings)，或[试用 Svelte 演示](/demo-svelte/)。
