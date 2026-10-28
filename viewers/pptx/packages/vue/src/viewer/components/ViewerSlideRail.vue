<script setup lang="ts">
/**
 * ViewerSlideRail: the desktop left-hand slide rail of `PowerPointViewer`.
 *
 * The rail itself is the public `SlidesPane` building block; this adapter
 * maps the viewer's slide and section operations onto its events and honours
 * a host `regions.slidesPane` replacement. It is handed the MERGED slides
 * (template layer folded in) so the rail matches what the canvas paints.
 *
 * Hidden on mobile by the parent, where it would otherwise collapse the slide
 * canvas to zero height; a phone navigates slides from the bottom bar instead.
 */
import { EDITOR_THUMBNAIL_WIDTH } from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxHeaderFooter,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';
import { computed } from 'vue';

import type { UseSectionOperationsResult } from '../composables/useSectionOperations';
import type { UseSlideOperationsResult } from '../composables/useSlideOperations';
import { useRegion } from '../composables/useViewerRegions';
import type { CanvasSize } from '../types';
import SlidesPane from './SlidesPane.vue';
import ViewerRegion from './ViewerRegion.vue';

/**
 * Shared preview width inside the 180px rail, leaving room for slide numbers.
 */
const THUMB_WIDTH = EDITOR_THUMBNAIL_WIDTH;

const props = defineProps<{
	/** Slides with the template (master/layout) layer merged in. */
	mergedSlides: PptxSlide[];
	activeSlideIndex: number;
	canvasSize: CanvasSize;
	mediaDataUrls: Map<string, string>;
	canEdit: boolean;
	hasSections: boolean;
	sectionOps: UseSectionOperationsResult;
	slideOps: UseSlideOperationsResult;
	goTo: (index: number) => void;
	toggleSlideHidden: (index: number) => void;
	/** Makes `index` active, then opens the Layout gallery anchored at (x, y). */
	onOpenLayoutForSlide: (index: number, x: number, y: number) => void;
	/** Deck header/footer settings, for the thumbnails' fields. */
	headerFooter?: PptxHeaderFooter;
	/** Custom document properties, for the thumbnails' `docproperty` fields. */
	customProperties?: readonly PptxCustomProperty[];
	/** `p:presentation/@firstSlideNum`, so thumbnails number slides as the canvas does. */
	firstSlideNumber?: number;
	/** Presentation theme, for tables in thumbnails. */
	theme?: PptxTheme;
	/** Parsed table styles, for tables in thumbnails. */
	tableStyleMap?: ParsedTableStyleMap;
}>();

// A host `regions.slidesPane` replaces the thumbnails; the rail keeps its width.
const region = useRegion('slidesPane');
const regionContext = computed(() => region.context(() => undefined));

// The section headers show only when the deck declares sections.
const sectionGroups = computed(() =>
	props.hasSections ? props.sectionOps.slidesBySection.value : [],
);
</script>

<template>
	<aside
		v-if="region.renderer.value && regionContext"
		data-pptx-chrome="slides"
		class="flex h-full shrink-0 flex-col border-r border-border bg-secondary/30"
		:style="{ width: `${THUMB_WIDTH + 48}px` }"
	>
		<ViewerRegion :renderer="region.renderer.value" :context="regionContext" />
	</aside>
	<SlidesPane
		v-else
		:slides="mergedSlides"
		:active-slide-index="activeSlideIndex"
		:canvas-size="canvasSize"
		:media-data-urls="mediaDataUrls"
		:section-groups="sectionGroups"
		:can-edit="canEdit"
		:header-footer="headerFooter"
		:custom-properties="customProperties"
		:first-slide-number="firstSlideNumber"
		:theme="theme"
		:table-style-map="tableStyleMap"
		@select="goTo"
		@move-slide="slideOps.moveSlide"
		@add-slide="slideOps.addSlide()"
		@add-slide-after="(i) => slideOps.addSlide(i)"
		@duplicate-slide="slideOps.duplicateSlide"
		@delete-slide="slideOps.deleteSlide"
		@toggle-slide-hidden="toggleSlideHidden"
		@open-layout-gallery="(r) => onOpenLayoutForSlide(r.slideIndex, r.x, r.y)"
		@add-section="sectionOps.addSection"
		@rename-section="sectionOps.renameSection"
		@delete-section="sectionOps.deleteSection"
		@move-section-up="sectionOps.moveSectionUp"
		@move-section-down="sectionOps.moveSectionDown"
		@toggle-section-collapse="sectionOps.toggleSectionCollapse"
	/>
</template>
