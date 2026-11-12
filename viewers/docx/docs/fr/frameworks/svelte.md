# Svelte

Utilisez l'adaptateur Svelte comme composant et gérez les changements de modèle avec le callback d'événement.

```svelte
<script>
	import WordEditor from 'docx-svelte-viewer';
	let model = $state(initialDocument);
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

La racine du paquet est le composant (export par défaut). Les assistants JavaScript simples, `mountEditor`, `loadDocument` et les types partagés, se trouvent sous `docx-svelte-viewer/runtime`.

Consultez le [contrat complet des adaptateurs](/fr/bindings) ou [essayez la démo Svelte](/demo-svelte/).
