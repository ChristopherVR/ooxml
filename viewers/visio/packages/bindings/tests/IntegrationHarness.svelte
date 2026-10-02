<script lang="ts">
  import { untrack } from 'svelte';
  import VisioViewer from '../src/VisioViewer.svelte';
  import type { ViewerOptions, ViewerHandle } from '../src/common.js';
  let { initial }: { initial: ViewerOptions } = $props();
  let options = $state.raw<ViewerOptions>(untrack(() => initial));
  let viewer: ReturnType<typeof VisioViewer>;
  export function update(next: ViewerOptions) { options = next; }
  export function getHandle(): ViewerHandle { return viewer.getHandle(); }
</script>
<VisioViewer bind:this={viewer} {...options} />
