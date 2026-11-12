# Vue 3

`visio-vue-viewer` vous donne le `<visio-viewer>` partagé comme intégration native Vue 3.

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

La ref du composant expose le même handle et émet aussi les événements natifs en kebab-case. Pour les grands documents immuables, utilisez `shallowRef` ou `markRaw` : réenvelopper un document dans un nouveau proxy réactif change son identité et demande un remplacement externe.

Les propriétés (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) et le handle sont identiques dans chaque framework : voir le [contrat des liaisons](/fr/bindings) et l'[API du visualiseur](/fr/api). Pour composer votre propre interface à partir des éléments du visualiseur, voir les [blocs de construction](/fr/building-blocks) ; pour partager un dessin, voir [collaboration](/fr/collaboration). [Essayer la démo Vue 3](/demo-vue/){target="_self"}.
