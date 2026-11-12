# Vue

El adaptador de Vue 3 es un componente con las props compartidas y eventos en kebab-case.

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-vue-viewer';

const workbook = shallowRef(createWorkbook());
const editor = shallowRef();
</script>

<template>
	<SpreadsheetEditor
		ref="editor"
		:workbook="workbook"
		file-name="Budget.xlsx"
		@workbook-change="workbook = $event"
		@selection-change="(s) => console.log(s.ref)"
	/>
</template>
```

Use `shallowRef` para el libro: es un objeto plano grande y la reactividad profunda no aporta nada. La ref de plantilla expone el [handle](/es/bindings#handle).

Consulte el [contrato completo de los enlaces](/es/bindings) o [pruebe la demo de Vue](/demo-vue/).
