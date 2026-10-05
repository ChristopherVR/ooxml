<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import VisioViewer from '../../src/VisioViewer.svelte';
  import type { Workspace } from '../../../../demo/workspace.js';

  // The Svelte demo: the VisioViewer component, its handle and the document as $state.
  let { workspace }: { workspace: Workspace } = $props();
  // The workspace is fixed for the page's lifetime; only its first document seeds the state.
  let document = $state(untrack(() => workspace.initialDocument));
  let viewer: ReturnType<typeof VisioViewer> | undefined = $state();
  onMount(() => {
    if (viewer) workspace.attach(viewer.getHandle(), (next) => (document = next));
  });
</script>

<VisioViewer bind:this={viewer} {document} events={workspace.events} />
