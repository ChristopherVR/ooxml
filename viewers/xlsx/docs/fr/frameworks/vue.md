# Vue

L’adaptateur Vue 3 est un composant doté des props communes et d’événements en kebab-case.

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

Utilisez `shallowRef` pour le classeur : c’est un gros objet simple, et la réactivité profonde n’apporte rien. La référence de gabarit expose la [poignée](/fr/bindings#handle).

Consultez le [contrat complet des liaisons](/fr/bindings) ou [essayez la démo Vue](/demo-vue/).
