<script lang="ts">
 import { buildingBlock, type ViewerController } from '../../bindings/src/building-blocks';
 let { controller }: { controller: ViewerController } = $props();
 let host: HTMLDivElement;
 $effect(() => buildingBlock('mountDiagramInspector')(host, controller));
</script>
<div bind:this={host}></div>
