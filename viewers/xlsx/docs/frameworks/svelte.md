# Svelte

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `xlsx-svelte-viewer` package; it needs only `svelte` next to it.
:::

Use the Svelte 5 adapter as a component and handle workbook changes with the callback props.

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

The package root is the component (default export). The plain-JavaScript helpers (`mountEditor`, `loadWorkbook`, `createWorkbook` and the shared types) are under `xlsx-svelte-viewer/runtime`. Component exports: `load`, `newWorkbook`, `save`, `saveBytes`, `download`, `markClean`, `select`, `getSelection`, `setActiveSheet`, `isDirty` and `getElement`.

See the [complete binding contract](/bindings) or [try the Svelte demo](/demo-svelte/).
