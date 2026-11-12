# Svelte 5

`visio-svelte-viewer` stellt den gemeinsamen `<visio-viewer>` als native Svelte-5-Integration bereit.

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

Callbacks laufen über die Prop `events`. Die Komponente exportiert `load`, `fit` und `getHandle`. Verwenden Sie `$state.raw` für große, unveränderliche Dokumente.

Die Eigenschaften (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) und das Handle sind in jedem Framework identisch: siehe den [Binding-Vertrag](/de/bindings) und die [Viewer-API](/de/api). Um stattdessen eine eigene Oberfläche aus den Teilen des Viewers zusammenzusetzen, siehe die [Bausteine](/de/building-blocks); um eine Zeichnung zu teilen, siehe [Zusammenarbeit](/de/collaboration). [Die Svelte-5-Demo ausprobieren](/demo-svelte/){target="_self"}.
