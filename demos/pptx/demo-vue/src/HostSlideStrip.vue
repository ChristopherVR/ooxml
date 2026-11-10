<script setup lang="ts">
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxElement,
	PptxHeaderFooter,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';
/** A host-owned thumbnail strip, fed from the values the built-in slides pane takes. */
import { SlideThumbnail } from 'pptx-vue-viewer';

import {
	cssText,
	HOST_STRIP_LABEL,
	HOST_STRIP_STYLE,
	HOST_STRIP_THUMBNAIL_WIDTH,
	hostStripItemLabel,
	hostStripItemStyle,
} from '../../shared/host-thumbnail-strip';

defineProps<{
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
	activeSlideIndex: number;
	canvasSize: { width: number; height: number };
	mediaDataUrls: Map<string, string>;
	headerFooter?: PptxHeaderFooter | undefined;
	customProperties?: readonly PptxCustomProperty[] | undefined;
	firstSlideNumber?: number | undefined;
	theme?: PptxTheme | undefined;
	tableStyleMap?: ParsedTableStyleMap | undefined;
}>();
const emit = defineEmits<{ select: [index: number] }>();
</script>

<template>
	<nav :aria-label="HOST_STRIP_LABEL" data-host-slide-strip :style="cssText(HOST_STRIP_STYLE)">
		<button
			v-for="(slide, index) in slides"
			:key="slide.id"
			type="button"
			:aria-label="hostStripItemLabel(index)"
			:aria-current="index === activeSlideIndex ? 'true' : undefined"
			:style="cssText(hostStripItemStyle(index === activeSlideIndex))"
			@click="emit('select', index)"
		>
			<SlideThumbnail
				:slide="slide"
				:template-elements="templateElementsBySlideId[slide.id] ?? []"
				:canvas-size="canvasSize"
				:preview-width="HOST_STRIP_THUMBNAIL_WIDTH"
				:media-data-urls="mediaDataUrls"
				:header-footer="headerFooter"
				:custom-properties="customProperties"
				:first-slide-number="firstSlideNumber"
				:theme="theme"
				:table-style-map="tableStyleMap"
			/>
		</button>
	</nav>
</template>
