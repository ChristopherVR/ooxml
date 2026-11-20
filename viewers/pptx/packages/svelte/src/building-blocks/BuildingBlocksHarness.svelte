<script lang="ts">
	/** Test host: the three blocks over one `createViewerBuildingBlocks`, options read live. */
	import ExternalContentBar from '../viewer/components/ExternalContentBar.svelte';
	import BuildingBlocksCanvas from './BuildingBlocksCanvas.svelte';
	import BuildingBlocksSlidesPane from './BuildingBlocksSlidesPane.svelte';
	import BuildingBlocksToolbar from './BuildingBlocksToolbar.svelte';
	import { createViewerBuildingBlocks } from './create-viewer-building-blocks.svelte';
	import type { ViewerBuildingBlocks, ViewerBuildingBlocksOptions } from './types';

	const {
		getOptions,
		onblocks,
		include,
	}: {
		getOptions: () => ViewerBuildingBlocksOptions;
		onblocks: (blocks: ViewerBuildingBlocks) => void;
		include?: boolean;
	} = $props();

	const blocks = createViewerBuildingBlocks(getOptions);
	onblocks(blocks);
</script>

{#if blocks.externalContent.barVisible}
	<ExternalContentBar
		onenable={blocks.externalContent.enable}
		onclose={blocks.externalContent.close}
	/>
{/if}
<BuildingBlocksToolbar {blocks} />
{#if include !== false}<BuildingBlocksSlidesPane {blocks} />{/if}
<div style="height: 400px"><BuildingBlocksCanvas {blocks} /></div>
