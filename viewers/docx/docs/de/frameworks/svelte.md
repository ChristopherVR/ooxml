# Svelte

Verwenden Sie den Svelte-Adapter als Komponente und behandeln Sie Modelländerungen über den Ereignis-Callback.

```svelte
<script>
	import WordEditor from 'docx-svelte-viewer';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

Das Paket-Root ist die Komponente (Default-Export). Die reinen JavaScript-Helfer `mountEditor`, `loadDocument` und die gemeinsamen Typen finden Sie unter `docx-svelte-viewer/runtime`.

Siehe den [vollständigen Binding-Vertrag](/de/bindings) oder [probieren Sie die Svelte-Demo aus](/demo-svelte/).
