<script lang="ts">
 import type { Readable } from 'svelte/store';
 import type { ViewerController } from '../../bindings/src/building-blocks';
 import DiagramCanvas from './DiagramCanvas.svelte';
 import Inspector from './Inspector.svelte';
 let { owner }: { owner: Readable<{ mount(node: HTMLElement | null): void; controller: ViewerController }> } = $props();
</script>
<DiagramCanvas mount={$owner.mount} /><Inspector controller={$owner.controller} />
