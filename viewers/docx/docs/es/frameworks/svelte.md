# Svelte

Use el adaptador de Svelte como componente y gestione los cambios del modelo con la función de callback del evento.

```svelte
<script>
	import WordEditor from 'docx-svelte-viewer';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

La raíz del paquete es el componente (exportación por defecto). Los auxiliares en JavaScript plano, `mountEditor`, `loadDocument` y los tipos compartidos, se encuentran en `docx-svelte-viewer/runtime`.

Consulte el [contrato completo de los adaptadores](/es/bindings) o [pruebe la demo de Svelte](/demo-svelte/).
