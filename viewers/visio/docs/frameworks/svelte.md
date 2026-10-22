# Svelte 5

`visio-svelte-viewer` gives you the shared `<visio-viewer>` as a native Svelte 5 integration.

```bash
npm install visio-svelte-viewer svelte
```

```svelte
<script lang="ts">
	import VisioViewer from 'visio-svelte-viewer';
	import type { VisioDocument } from 'visio-svelte-viewer';
	let { diagram }: { diagram: VisioDocument } = $props();
</script>

<VisioViewer document={diagram} style="height: 600px" events={{ 'document-error': console.error }} />
```

Callbacks go through the `events` prop. The component exports `load`, `fit` and `getHandle`. Use `$state.raw` for large immutable documents.

The properties (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `ribbonAddIns`, `collaboration`, `events`) and the handle are identical in every framework: see the [binding contract](/bindings) and the [viewer API](/api). To compose your own UI from the viewer's parts instead, see the [building blocks](/building-blocks); to share a drawing, see [collaboration](/collaboration). [Try the Svelte 5 demo](/demo-svelte/){target="_self"}.
