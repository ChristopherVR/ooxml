# Svelte

将 Svelte 适配器作为组件使用，并通过事件回调处理模型变更。

```svelte
<script>
	import WordEditor from 'docx-svelte-viewer';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

包的根入口即为组件（默认导出）。纯 JavaScript 辅助函数 `mountEditor`、`loadDocument` 以及共享类型位于 `docx-svelte-viewer/runtime` 下。

参见[完整的绑定契约](/zh/bindings)或[试用 Svelte 演示](/demo-svelte/)。
