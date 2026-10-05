# Vue

The Vue 3 adapter is a component with the shared props and kebab-case events.

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

Use `shallowRef` for the workbook: it is a large plain object and deep reactivity adds nothing. The template ref exposes the [handle](/bindings#handle).

See the [complete binding contract](/bindings) or [try the Vue demo](/demo-vue/).
