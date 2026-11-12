# Vue 3

`visio-vue-viewer` stellt den gemeinsamen `<visio-viewer>` als native Vue-3-Integration bereit.

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

Die Komponenten-Ref stellt dasselbe Handle bereit und sendet außerdem die nativen Ereignisse in kebab-case. Verwenden Sie bei großen, unveränderlichen Dokumenten `shallowRef` oder `markRaw`: Ein Dokument in einen neuen reaktiven Proxy einzuwickeln ändert seine Identität und fordert ein externes Ersetzen an.

Die Eigenschaften (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) und das Handle sind in jedem Framework identisch: siehe den [Binding-Vertrag](/de/bindings) und die [Viewer-API](/de/api). Um stattdessen eine eigene Oberfläche aus den Teilen des Viewers zusammenzusetzen, siehe die [Bausteine](/de/building-blocks); um eine Zeichnung zu teilen, siehe [Zusammenarbeit](/de/collaboration). [Die Vue-3-Demo ausprobieren](/demo-vue/){target="_self"}.
