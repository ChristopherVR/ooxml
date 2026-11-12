# Vue

Der Vue-Adapter stellt den Editor als Komponente mit Modell-Props und Vue-Ereignissen bereit.

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
const model = defineModel();
</script>

<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Siehe den [vollständigen Binding-Vertrag](/de/bindings) oder [probieren Sie die Vue-Demo aus](/demo-vue/).
