# Vue

L'adaptateur Vue expose l'éditeur comme un composant avec des props de modèle et des événements Vue.

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
const model = defineModel();
</script>

<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Consultez le [contrat complet des adaptateurs](/fr/bindings) ou [essayez la démo Vue](/demo-vue/).
