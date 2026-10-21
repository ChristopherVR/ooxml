<script lang="ts">
 import { prepareDiagramBlock } from '../../bindings/src/building-blocks';
 function theme(host: HTMLElement) { prepareDiagramBlock(host, 'tabs'); }
 import type { DiagramToolbarProps } from '../../bindings/src/building-blocks';
 let { toolbarProps }: { toolbarProps: DiagramToolbarProps } = $props();
</script>
<select use:theme aria-label={toolbarProps.labels.Page} value={toolbarProps.pageIndex} onchange={event => toolbarProps.onPageChange(Number(event.currentTarget.value))}>
 {#each toolbarProps.pages as page, index (page.id)}<option value={index}>{page.name}</option>{/each}
</select>
