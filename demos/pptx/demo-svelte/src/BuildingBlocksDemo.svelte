<script lang="ts">
	import { BuildingBlocksCanvas, BuildingBlocksSlidesPane, BuildingBlocksToolbar, createViewerBuildingBlocks } from 'pptx-svelte-viewer';
	import type { ViewerBuildingBlocksOptions } from 'pptx-svelte-viewer';
	import { onMount } from 'svelte';

	import { mountBuildingBlocksPanel } from '../../shared/building-blocks-panel';

	/**
	 * `?blocks=1`: a custom shell composed from the Svelte building blocks (toolbar, slides pane,
	 * canvas) with the shared host panel that finishes their workflows. The framework-neutral
	 * `e2e/pptx/building-blocks.spec.ts` drives it.
	 */
	let options = $state.raw<ViewerBuildingBlocksOptions>({ canEdit: false, fileName: 'sample-deck.pptx' });
	const blocks = createViewerBuildingBlocks(() => options);
	let panelHost = $state<HTMLElement>();
	let panel = $state.raw<{ refresh(): void }>();

	onMount(() => {
		void fetch(`${import.meta.env.BASE_URL}sample-deck.pptx`)
			.then((response) => response.arrayBuffer())
			.then((bytes) => {
				options = { ...options, content: new Uint8Array(bytes) };
			});
	});
	$effect(() => {
		if (panelHost && !panel) {
			panel = mountBuildingBlocksPanel(panelHost, {
				actions: () => blocks.actions,
				state: () => blocks.state,
				setCanEdit: (canEdit) => {
					options = { ...options, canEdit };
				},
				setCollaboration: (collaboration) => {
					options = {
						...options,
						collaboration: collaboration as ViewerBuildingBlocksOptions['collaboration'],
					};
				},
			});
		}
	});
	// Re-render the host panel whenever anything it reads changes.
	$effect(() => {
		panel?.refresh();
	});
</script>

<section bind:this={panelHost}></section>
<BuildingBlocksToolbar {blocks} />
<div class="body">
	<div class="rail"><BuildingBlocksSlidesPane {blocks} /></div>
	<div class="canvas"><BuildingBlocksCanvas {blocks} /></div>
</div>

<style>
	.body { display: flex; height: calc(100vh - 160px); }
	.rail { width: 200px; overflow: auto; }
	.canvas { flex: 1; position: relative; min-width: 0; }
</style>
