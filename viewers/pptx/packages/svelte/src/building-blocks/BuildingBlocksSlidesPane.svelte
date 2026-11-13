<script lang="ts">
	/**
	 * BuildingBlocksSlidesPane: the slide thumbnail rail block, fed from the viewer state the
	 * building blocks made (`blocks.slidesPaneProps`), so its edits share the canvas's history.
	 */
	import SlidesPane from '../viewer/components/SlidesPane.svelte';
	import type { ViewerBuildingBlocks } from './types';

	const { blocks, class: className = '' }: { blocks: ViewerBuildingBlocks; class?: string } =
		$props();
</script>

<div class={`pptx-svelte-viewer pptx-svelte-bb-slides-pane ${className}`} data-pptx-block="slides-pane">
	<SlidesPane {...blocks.slidesPaneProps} />
</div>

<style>
	.pptx-svelte-bb-slides-pane {
		display: flex;
		flex-direction: column;
		height: 100%;
		min-height: 0;
		overflow: auto;
	}
</style>
