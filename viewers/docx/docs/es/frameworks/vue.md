# Vue

El adaptador de Vue expone el editor como un componente con propiedades de modelo y eventos de Vue.

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
const model = defineModel();
</script>

<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Consulte el [contrato completo de los adaptadores](/es/bindings) o [pruebe la demo de Vue](/demo-vue/).
