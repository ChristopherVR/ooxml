<script lang="ts">
 import { buildingBlock, type ViewerController } from '../../bindings/src/building-blocks';
 let { controller, locale }: { controller: ViewerController; locale?: string | undefined } = $props();
 let host: HTMLDivElement;
 $effect(() => buildingBlock('mountDiagramShapes')(host, controller, { locale }));
</script>
<div bind:this={host}></div>
