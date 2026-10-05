<script lang="ts">
  import { untrack } from 'svelte';
  import VisioViewer from '../src/VisioViewer.svelte';
  import type { ViewerCallbacks, ViewerHandle } from '../src/common.js';
  let { initialEvents }: { initialEvents: ViewerCallbacks } = $props();
  let zoom = $state(2);
  let events = $state<ViewerCallbacks | undefined>(untrack(() => initialEvents));
  let viewer: ReturnType<typeof VisioViewer>;
  export function update(next: number, callbacks?: ViewerCallbacks) { zoom = next; events = callbacks; }
  export function getHandle(): ViewerHandle { return viewer.getHandle(); }
  export function getViewer() { return viewer; }
</script>
<VisioViewer bind:this={viewer} document={null} pageIndex={0} {zoom} showToolbar={false} {events} />
