<script setup lang="ts">
/**
 * BuildingBlocksCanvas: the slide canvas block. It runs the full `PowerPointViewer` with its chrome
 * turned off (no ribbon, rail, status bar, inspector, notes, title bar, toasts or dialogs); the
 * viewer's own dialogs become host requests through the bridge it is given.
 */
import { computed, onBeforeUnmount, onMounted, provide, ref, watch } from 'vue';

import PowerPointViewer from '../viewer/PowerPointViewer.vue';
import type { PowerPointViewerExpose } from '../viewer/types';
import { BuildingBlocksBridgeKey } from './bridge';
import { syncFindHighlights } from './find-highlight';
import type { ViewerBuildingBlocks } from './types';

const props = defineProps<{ blocks: ViewerBuildingBlocks }>();

/** The regions the canvas block leaves out: the host composes its own shell around it. */
const BLOCK_HIDDEN_PANELS = [
	'statusBar',
	'slidesPane',
	'inspector',
	'notes',
	'quickAccessToolbar',
	'titleBar',
	'selectionPane',
	'accessibilityPane',
	'comparePane',
] as const;

provide(BuildingBlocksBridgeKey, props.blocks.bridge);

const root = ref<HTMLElement | null>(null);
const viewer = ref<PowerPointViewerExpose | null>(null);
watch(viewer, (handle) => {
	props.blocks.bridge.viewer.value = handle;
});

const o = computed(() => props.blocks.options.value);
const customization = computed(() => ({
	...o.value.customization,
	hiddenPanels: [...(o.value.customization?.hiddenPanels ?? []), ...BLOCK_HIDDEN_PANELS],
}));
// Everything the blocks decide or forward as an event is left out of the plain props.
const viewerProps = computed(() => {
	const {
		content: _content,
		canEdit: _canEdit,
		customization: _customization,
		onRequest: _onRequest,
		fileLimits: _fileLimits,
		onContentChange: _a,
		onDirtyChange: _b,
		onActiveSlideChange: _c,
		onModeChange: _d,
		onZoomChange: _e,
		onSelectionChange: _f,
		onSlideCountChange: _g,
		onOpenLayoutGallery: _h,
		...rest
	} = o.value;
	return rest;
});

// The stage is repainted on change: outline the find matches again after every paint.
let observer: MutationObserver | undefined;
const highlight = () => {
	const engine = props.blocks.bridge.engine.value;
	if (root.value && engine) {
		const { matches, current, slideIndex } = engine.findHighlights();
		syncFindHighlights(root.value, matches, current, slideIndex);
	}
};
onMounted(() => {
	if (root.value) {
		observer = new MutationObserver(highlight);
		observer.observe(root.value, { childList: true, subtree: true });
	}
});
onBeforeUnmount(() => {
	observer?.disconnect();
	props.blocks.bridge.viewer.value = null;
});
watch(() => props.blocks.bridge.engine.value?.findHighlights(), highlight, { flush: 'post' });
</script>

<template>
	<div ref="root" data-pptx-block="canvas" class="pptx-vue-bb-canvas">
		<PowerPointViewer
			ref="viewer"
			v-bind="viewerProps"
			:content="o.content ?? (null as unknown as Uint8Array)"
			:can-edit="o.canEdit === true"
			:show-toolbar="false"
			:show-thumbnails="false"
			:show-compatibility-toasts="false"
			:customization="customization"
			@content-change="o.onContentChange?.($event)"
			@dirty-change="o.onDirtyChange?.($event)"
			@active-slide-change="o.onActiveSlideChange?.($event)"
			@mode-change="o.onModeChange?.($event)"
			@zoom-change="o.onZoomChange?.($event)"
			@selection-change="o.onSelectionChange?.($event)"
			@slide-count-change="o.onSlideCountChange?.($event)"
		/>
	</div>
</template>

<style>
.pptx-vue-bb-canvas {
	position: relative;
	width: 100%;
	height: 100%;
}
.pptx-vue-bb-canvas [data-pptx-find-match] {
	outline: 2px dashed var(--pptx-primary, #d24726);
	outline-offset: 2px;
}
.pptx-vue-bb-canvas [data-pptx-find-match='current'] {
	outline-style: solid;
}
</style>
