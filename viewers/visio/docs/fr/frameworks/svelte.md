# Svelte 5

`visio-svelte-viewer` vous donne le `<visio-viewer>` partagé comme intégration native Svelte 5.

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

Les rappels passent par la prop `events`. Le composant exporte `load`, `fit` et `getHandle`. Utilisez `$state.raw` pour les grands documents immuables.

Les propriétés (`document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `fileProviders`, `customization`, `events`) et le handle sont identiques dans chaque framework : voir le [contrat des liaisons](/fr/bindings) et l'[API du visualiseur](/fr/api). Pour composer votre propre interface à partir des éléments du visualiseur, voir les [blocs de construction](/fr/building-blocks) ; pour partager un dessin, voir [collaboration](/fr/collaboration). [Essayer la démo Svelte 5](/demo-svelte/){target="_self"}.
