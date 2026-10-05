# Vue 3

`visio-vue-viewer` gives you the shared `<visio-viewer>` as a native Vue 3 integration.

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

The component ref exposes the same handle and also emits the native kebab-case events. For large immutable documents use `shallowRef` or `markRaw`: rewrapping a document in a new reactive proxy changes its identity and requests an external replacement.

The properties (`document`, `pageIndex`, `zoom`, `showToolbar`, `events`) and the handle are identical in every framework: see the [binding contract](/bindings) and the [viewer API](/api). [Try the Vue 3 demo](/demo-vue/){target="_self"}.
