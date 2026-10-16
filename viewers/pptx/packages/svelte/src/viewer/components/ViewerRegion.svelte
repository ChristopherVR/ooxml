<script lang="ts">
	/**
	 * ViewerRegion: the one host-supplied region container. Builds the region
	 * context from the root's base context plus the site's `region`, `close`
	 * and `extra` (the Options model), and hands it to `ViewerRegionHost`
	 * through the `viewerRegionHost` action.
	 */
	import type { ViewerRegionContext, ViewerRegionRenderer } from 'ooxml-ui/pptx';

	import { useRegionBase, viewerRegionHost } from '../state/viewer-regions';

	const {
		region,
		renderer,
		close = () => undefined,
		extra,
	}: {
		region: ViewerRegionContext['region'];
		renderer: ViewerRegionRenderer;
		close?: () => void;
		extra?: Record<string, unknown>;
	} = $props();

	const getBase = useRegionBase();
	const context = $derived.by(() => {
		const base = getBase?.();
		return base ? ({ ...base, ...extra, region, close } as ViewerRegionContext) : undefined;
	});
</script>

{#if context}
	<div class="pptx-svelte-region" use:viewerRegionHost={{ renderer, context }}></div>
{/if}

<style>
	.pptx-svelte-region {
		display: flex;
		flex-direction: column;
		flex: 1;
		width: 100%;
		height: 100%;
		min-width: 0;
		min-height: 0;
	}
</style>
