# Svelte 5

`visio-svelte-viewer` 将共享的 `<visio-viewer>` 以原生 Svelte 5 集成的形式提供给你。

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

回调通过 `events` 属性传入。该组件导出 `load`、`fit` 和 `getHandle`。对于大型的不可变文档，请使用 `$state.raw`。

各框架中的属性（`document`、`pageIndex`、`zoom`、`showToolbar`、`canEdit`、`readOnly`、`ribbonAddIns`、`collaboration`、`storage`、`fileProviders`、`customization`、`events`）和句柄完全相同：请参阅[绑定约定](/zh/bindings)和[查看器 API](/zh/api)。如需改用查看器的各个部件组合自己的界面，请参阅[构建模块](/zh/building-blocks)；如需共享图纸，请参阅[协作](/zh/collaboration)。[试用 Svelte 5 演示](/demo-svelte/){target="_self"}。
