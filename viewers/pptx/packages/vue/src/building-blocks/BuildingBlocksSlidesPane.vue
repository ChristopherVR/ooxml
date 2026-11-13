<script setup lang="ts">
/**
 * BuildingBlocksSlidesPane: the slide thumbnail rail block, fed from the viewer behind the canvas
 * block (same deck, active slide and slide and section commands).
 */
import { computed } from 'vue';

import SlidesPane from '../viewer/components/SlidesPane.vue';
import { provideElementRenderErrorReporter } from '../viewer/composables/element-render-errors';
import type { ViewerBuildingBlocks } from './types';

const props = defineProps<{ blocks: ViewerBuildingBlocks }>();
const pane = computed(() => props.blocks.bridge.engine.value?.slidesPane() ?? null);
// The rail renders outside the canvas block's viewer, so it takes the reporter itself.
provideElementRenderErrorReporter(() => props.blocks.options.value.onElementRenderError);

function openLayoutGallery(request: { slideIndex: number; x: number; y: number }): void {
	props.blocks.options.value.onOpenLayoutGallery?.(request);
}
</script>

<template>
	<div data-pptx-block="slides-pane" class="pptx-vue-viewer pptx-vue-bb-slides-pane">
		<SlidesPane
			v-if="pane"
			:slides="[...pane.slides]"
			:active-slide-index="pane.activeSlideIndex"
			:canvas-size="pane.canvasSize"
			:media-data-urls="pane.mediaDataUrls"
			:section-groups="pane.sectionGroups"
			:can-edit="pane.canEdit"
			:header-footer="pane.headerFooter"
			:custom-properties="pane.customProperties"
			:first-slide-number="pane.firstSlideNumber"
			:theme="pane.theme"
			:table-style-map="pane.tableStyleMap"
			@select="pane.goTo"
			@move-slide="pane.slideOps.moveSlide"
			@add-slide="pane.slideOps.addSlide()"
			@add-slide-after="pane.slideOps.addSlide"
			@duplicate-slide="pane.slideOps.duplicateSlide"
			@delete-slide="pane.slideOps.deleteSlide"
			@toggle-slide-hidden="pane.toggleSlideHidden"
			@open-layout-gallery="openLayoutGallery"
			@add-section="pane.sectionOps.addSection"
			@rename-section="pane.sectionOps.renameSection"
			@delete-section="pane.sectionOps.deleteSection"
			@move-section-up="pane.sectionOps.moveSectionUp"
			@move-section-down="pane.sectionOps.moveSectionDown"
			@toggle-section-collapse="pane.sectionOps.toggleSectionCollapse"
		/>
	</div>
</template>
