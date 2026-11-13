<script lang="ts">
	/** Test host: the three blocks over one `createViewerBuildingBlocks`, options read live. */
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

<BuildingBlocksToolbar {blocks} />
{#if include !== false}<BuildingBlocksSlidesPane {blocks} />{/if}
<div style="height: 400px"><BuildingBlocksCanvas {blocks} /></div>
