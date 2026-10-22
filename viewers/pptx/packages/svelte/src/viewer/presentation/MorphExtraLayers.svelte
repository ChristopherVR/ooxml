<script lang="ts">
	/**
	 * The morph-only layers `PresentationTransitionOverlay.svelte` paints ABOVE
	 * its incoming layer, split out to keep that file under the project's
	 * per-file LOC budget:
	 *
	 *  - the **overlay layers**: the plan's ghosts and lifted arrivals, one
	 *    layer per run in the plan's single back-to-front order (issue #141).
	 *    Every layer is BACKGROUND-FREE, or its slide fill would cover the morph
	 *    beneath it; a lifted run holds arriving shapes a ghost would otherwise
	 *    hide (issue #146), their copy on the incoming layer held invisible.
	 *  - the **crossfade groups**: pairs the overlay paints BOTH halves of, each
	 *    as one isolated group so the halves are SUMMED rather than stacked
	 *    (issue #161).
	 */
	import type { CanvasSize } from 'ooxml-ui/pptx';

	import type {
		MorphCrossfadeGroupView,
		MorphOverlayLayerView,
	} from './use-morph-transition-overlay.svelte';
	import SlideStage from '../components/SlideStage.svelte';

	const {
		overlayLayers,
		crossfadeGroups,
		canvasSize,
		mediaDataUrls,
		scale = 1,
	}: {
		overlayLayers: readonly MorphOverlayLayerView[];
		crossfadeGroups: readonly MorphCrossfadeGroupView[];
		canvasSize: CanvasSize;
		mediaDataUrls: Map<string, string>;
		scale?: number;
	} = $props();
</script>

{#each overlayLayers as layer (layer.key)}
	<div
		class="pptx-svelte-transition-layer"
		data-pptx-transition-layer={layer.side === 'outgoing' ? 'outgoing' : 'lifted'}
		data-pptx-morph-outgoing={layer.side === 'outgoing' ? 'true' : undefined}
		data-pptx-morph-lifted={layer.side === 'incoming' ? 'true' : undefined}
		style={layer.style}
	>
		<SlideStage slide={layer.slide} {canvasSize} {mediaDataUrls} {scale} transparentBackground />
	</div>
{/each}
{#each crossfadeGroups as group (group.key)}
	<div data-pptx-morph-crossfade={group.key} style={group.style}>
		<div
			class="pptx-svelte-transition-layer"
			data-pptx-transition-layer="outgoing"
			data-pptx-morph-outgoing="true"
			style={group.outgoingStyle}
		>
			<SlideStage slide={group.outgoing} {canvasSize} {mediaDataUrls} {scale} transparentBackground />
		</div>
		<div
			class="pptx-svelte-transition-layer"
			data-pptx-transition-layer="lifted"
			data-pptx-morph-lifted="true"
			style={group.incomingStyle}
		>
			<SlideStage slide={group.incoming} {canvasSize} {mediaDataUrls} {scale} transparentBackground />
		</div>
	</div>
{/each}

<style>
	/* Same rule as the parent overlay's own `.pptx-svelte-transition-layer`. */
	.pptx-svelte-transition-layer {
		position: absolute;
		inset: 0;
		overflow: hidden;
		will-change: transform, opacity, clip-path, filter;
	}
</style>
