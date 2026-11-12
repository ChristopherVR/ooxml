# Svelte 5

`visio-svelte-viewer` le ofrece el `<visio-viewer>` compartido como integración nativa de Svelte 5.

```bash
npm install visio-svelte-viewer svelte
```

```svelte
<script lang="ts">
	import VisioViewer from 'visio-svelte-viewer';
	import type { VisioDocument } from 'visio-svelte-viewer';
	let { diagram }: { diagram: VisioDocument } = $props();
</script>

<VisioViewer document={diagram} style="height: 600px" events={{ 'document-error': console.error }} />
```

Los callbacks pasan por la prop `events`. El componente exporta `load`, `fit` y `getHandle`. Use `$state.raw` para documentos grandes e inmutables.

Las propiedades (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) y el handle son idénticos en todos los frameworks: consulte el [contrato de enlaces](/es/bindings) y la [API del visor](/es/api). Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/es/building-blocks); para compartir un dibujo, consulte [colaboración](/es/collaboration). [Pruebe la demostración de Svelte 5](/demo-svelte/){target="_self"}.
