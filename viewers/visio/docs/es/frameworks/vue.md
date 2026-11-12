# Vue 3

`visio-vue-viewer` le ofrece el `<visio-viewer>` compartido como integración nativa de Vue 3.

```bash
npm install visio-vue-viewer vue
```

```vue
<script setup lang="ts">
import { VisioViewer } from 'visio-vue-viewer';
import type { VisioDocument } from 'visio-vue-viewer';
defineProps<{ diagram: VisioDocument }>();
</script>

<template>
	<VisioViewer :document="diagram" style="height: 600px" @document-error="console.error" />
</template>
```

La ref del componente expone el mismo handle y además emite los eventos nativos en kebab-case. Para documentos grandes e inmutables use `shallowRef` o `markRaw`: envolver un documento en un nuevo proxy reactivo cambia su identidad y solicita una sustitución externa.

Las propiedades (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) y el handle son idénticos en todos los frameworks: consulte el [contrato de enlaces](/es/bindings) y la [API del visor](/es/api). Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/es/building-blocks); para compartir un dibujo, consulte [colaboración](/es/collaboration). [Pruebe la demostración de Vue 3](/demo-vue/){target="_self"}.
