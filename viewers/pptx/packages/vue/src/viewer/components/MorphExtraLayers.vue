<script setup lang="ts">
import type {
	MorphCrossfadeGroupView,
	MorphOverlayLayerView,
} from '../composables/use-morph-transition-overlay';
import type { CanvasSize } from '../types';
import SlideStage from './SlideStage.vue';

/**
 * The morph-only layers `PresentationTransitionOverlay.vue` paints ABOVE its
 * incoming layer, split out to keep that file under the project's per-file
 * LOC budget:
 *
 *  - the **overlay layers**: the plan's ghosts and lifted arrivals, one layer
 *    per run in the plan's single back-to-front order (issue #141). A ghost
 *    run is BACKGROUND-FREE, or its slide fill would cover the morph beneath
 *    it; a lifted run holds arriving shapes a ghost would otherwise hide
 *    (issue #146), their copy on the incoming layer held invisible.
 *  - the **crossfade groups**: pairs the overlay paints BOTH halves of, each
 *    as one isolated group so the halves are SUMMED rather than stacked
 *    (issue #161).
 */
defineProps<{
	overlayLayers: readonly MorphOverlayLayerView[];
	crossfadeGroups: readonly MorphCrossfadeGroupView[];
	canvasSize: CanvasSize;
	mediaDataUrls: Map<string, string>;
	scale: number;
}>();
</script>

<template>
	<div
		v-for="layer in overlayLayers"
		:key="layer.key"
		class="pptx-vue-transition-layer"
		:data-pptx-transition-layer="layer.side === 'outgoing' ? 'outgoing' : 'lifted'"
		:data-pptx-morph-outgoing="layer.side === 'outgoing' ? 'true' : undefined"
		:data-pptx-morph-lifted="layer.side === 'incoming' ? 'true' : undefined"
		:style="layer.style"
	>
		<SlideStage
			:slide="layer.slide"
			:canvas-size="canvasSize"
			:media-data-urls="mediaDataUrls"
			:scale="scale"
			preserve-element-ids
			transparent-background
		/>
	</div>

	<div
		v-for="group in crossfadeGroups"
		:key="group.key"
		:data-pptx-morph-crossfade="group.key"
		:style="group.style"
	>
		<div
			class="pptx-vue-transition-layer"
			data-pptx-transition-layer="outgoing"
			data-pptx-morph-outgoing="true"
			:style="group.outgoingStyle"
		>
			<SlideStage
				:slide="group.outgoingSlide"
				:canvas-size="canvasSize"
				:media-data-urls="mediaDataUrls"
				:scale="scale"
				preserve-element-ids
				transparent-background
			/>
		</div>
		<div
			class="pptx-vue-transition-layer"
			data-pptx-transition-layer="lifted"
			data-pptx-morph-lifted="true"
			:style="group.incomingStyle"
		>
			<SlideStage
				:slide="group.incomingSlide"
				:canvas-size="canvasSize"
				:media-data-urls="mediaDataUrls"
				:scale="scale"
				preserve-element-ids
				transparent-background
			/>
		</div>
	</div>
</template>

<style scoped>
/* Same rule as the parent overlay's own `.pptx-vue-transition-layer`: the
   layer must FILL its ancestor, not shrink-wrap `SlideStage` (which scales
   with a CSS `transform` that never changes its laid-out box). */
.pptx-vue-transition-layer {
	position: absolute;
	inset: 0;
	overflow: hidden;
	will-change: transform, opacity, clip-path, filter;
}
</style>
