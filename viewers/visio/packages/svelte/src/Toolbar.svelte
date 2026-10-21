<script lang="ts">
 import { prepareDiagramBlock } from '../../bindings/src/building-blocks';
 function theme(host: HTMLElement) { prepareDiagramBlock(host, 'toolbar'); }
 import type { DiagramToolbarProps } from '../../bindings/src/building-blocks';
 let { toolbarProps }: { toolbarProps: DiagramToolbarProps } = $props();
</script>
<div role="toolbar" use:theme>
 <button type="button" disabled={!toolbarProps.canUndo} onclick={() => void toolbarProps.onUndo()}>{toolbarProps.labels.Undo}</button>
 <button type="button" disabled={!toolbarProps.canRedo} onclick={() => void toolbarProps.onRedo()}>{toolbarProps.labels.Redo}</button>
 <button type="button" onclick={toolbarProps.onFit}>{toolbarProps.labels['Fit to Window']}</button>
</div>
