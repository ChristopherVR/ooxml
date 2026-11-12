# Vue

Der Vue-3-Adapter ist eine Komponente mit den gemeinsamen Props und Ereignissen in Kebab-Case.

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

Verwenden Sie `shallowRef` für die Arbeitsmappe: Sie ist ein großes, einfaches Objekt, und tiefe Reaktivität bringt nichts. Die Template-Referenz stellt das [Handle](/de/bindings#handle) bereit.

Siehe die [vollständige Binding-Vertragsbeschreibung](/de/bindings) oder [probieren Sie die Vue-Demo aus](/demo-vue/).
