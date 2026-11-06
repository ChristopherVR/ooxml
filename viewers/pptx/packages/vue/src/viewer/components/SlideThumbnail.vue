<script setup lang="ts">
/**
 * SlideThumbnail: one slide drawn the way `SlidesPane` draws it, for a host that keeps its own
 * thumbnail strip.
 *
 * It paints the slide's template (layout/master) elements beneath its own, with the slide
 * background, through the same static `SlideStage` the rail uses, scaled to `previewWidth`.
 * The stage is static: no selection, editing, pointer handlers, focus stops or element ids, so
 * it never competes with the editing canvas. Fields and table styles resolve from the same deck
 * props `SlidesPane` takes, so a host feeds both from one place.
 *
 * A slide with a transition (other than Cut) shows its name in a corner badge, from the shared
 * `slideThumbnailTransitionBadge`, as every binding's thumbnail does; `SlidesPane` draws none.
 */
import {
	buildPreviewElements,
	EDITOR_THUMBNAIL_WIDTH,
	slideThumbnailLayout,
	slideThumbnailTransitionBadge,
	slideThumbnailTransitionBadgeStyle,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { provideSlidesPaneContexts } from '../composables/slides-pane-contexts';
import type { SlideThumbnailProps } from './slide-thumbnail-types';
import SlideStage from './SlideStage.vue';

const props = withDefaults(defineProps<SlideThumbnailProps>(), {
	templateElements: () => [] as PptxElement[],
	previewWidth: EDITOR_THUMBNAIL_WIDTH,
	mediaDataUrls: () => new Map<string, string>(),
	headerFooter: undefined,
	customProperties: undefined,
	firstSlideNumber: undefined,
	locale: undefined,
	theme: undefined,
	tableStyleMap: undefined,
});

provideSlidesPaneContexts(() => props);

const { t } = useI18n();

const stageSize = computed(() => ({
	width: Math.max(1, props.canvasSize.width),
	height: Math.max(1, props.canvasSize.height),
}));
const layout = computed(() =>
	slideThumbnailLayout(stageSize.value.width, stageSize.value.height, props.previewWidth),
);
// Template first, slide elements on top. `buildPreviewElements` honours "Hide Background
// Graphics", so the stage gets no separate template layer.
const previewSlide = computed<PptxSlide>(() => ({
	...props.slide,
	elements: buildPreviewElements(props.slide, props.templateElements),
}));
const transitionBadge = computed(() =>
	slideThumbnailTransitionBadge(props.slide, (key, params) => t(key, params ?? {})),
);
</script>

<template>
	<div
		class="relative w-full overflow-hidden bg-white"
		:style="{ height: `${layout.height}px`, overflow: 'hidden' }"
	>
		<SlideStage
			:slide="previewSlide"
			:canvas-size="stageSize"
			:media-data-urls="mediaDataUrls"
			:scale="layout.scale"
		/>
		<div
			v-if="transitionBadge"
			data-pptx-thumbnail-transition
			:title="transitionBadge.title"
			:style="slideThumbnailTransitionBadgeStyle(layout)"
		>
			{{ transitionBadge.label }}
		</div>
	</div>
</template>
