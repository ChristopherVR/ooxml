<script setup lang="ts">
/** Slide previews over the shared `createSlideSorterController`; this file only draws it. */
import { X } from '@lucide/vue';
import {
	createSlideSorterController,
	HIDDEN_SLIDE_SLASH_GRADIENT,
	hiddenSlideCue,
	slideClipboardLabels,
	slideSorterContextMenuLabel,
	slideSorterView,
} from 'ooxml-ui/pptx';
import type { SlideSorterContextMenuCommandId, SlidePasteRequest } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import type { CSSProperties } from 'vue';
import { computed, onBeforeUnmount, onMounted } from 'vue';
import { useI18n } from 'vue-i18n';

import { useController } from '../composables/useController';
import type { CanvasSize } from '../types';
import ContextMenu from './ContextMenu.vue';
import type { ContextMenuItem } from './ContextMenu.vue';
import SlideStage from './SlideStage.vue';

const { t } = useI18n();

const props = defineProps<{
	slides: PptxSlide[];
	clipboardSourceSlides?: PptxSlide[];
	clipboardTemplates?: Record<string, PptxElement[]>;
	canvasSize: CanvasSize;
	mediaDataUrls: Map<string, string>;
	activeIndex: number;
	canEdit?: boolean;
	copySystemSlides?: (indexes: number[]) => Promise<void>;
	pasteSystemSlides?: (insertAt: number) => Promise<void>;
}>();

const emit = defineEmits<{
	select: [index: number];
	reorder: [from: number, to: number];
	duplicate: [index: number];
	paste: [request: SlidePasteRequest];
	delete: [index: number];
	'toggle-hidden': [index: number];
	close: [];
}>();

const controller = createSlideSorterController({
	slides: () => props.slides,
	activeIndex: () => props.activeIndex,
	canEdit: () => Boolean(props.canEdit),
	clipboardSourceSlides: () => props.clipboardSourceSlides,
	clipboardTemplates: () => props.clipboardTemplates,
	onOperation: (operation, indexes) => {
		for (const index of indexes) {
			if (operation === 'delete') {
				emit('delete', index);
			} else if (operation === 'duplicate') {
				emit('duplicate', index);
			} else {
				emit('toggle-hidden', index);
			}
		}
	},
	onPaste: (request) => emit('paste', request),
	onMove: (from, to) => emit('reorder', from, to),
	onClose: () => emit('close'),
});
const state = useController(controller);
const view = computed(() => slideSorterView(state.value, props.slides));

/** Fixed thumbnail width (px); height derives from the canvas aspect ratio. */
const tileWidth = computed(() => 192 * view.value.zoomScale);
const tileScale = computed(() => tileWidth.value / Math.max(1, props.canvasSize.width));
const tileHeight = computed(() => Math.round(props.canvasSize.height * tileScale.value));
const stageWrapStyle = computed<CSSProperties>(() => ({
	width: `${tileWidth.value}px`,
	height: `${tileHeight.value}px`,
}));

/** The shared rail/sorter cue: the slash across the number and a description. */
const hiddenCue = hiddenSlideCue;
const slashGradient = HIDDEN_SLIDE_SLASH_GRADIENT;

const zoom = computed({
	get: () => state.value.sorter.zoom,
	set: (value: number) => controller.setZoom(value),
});

function onDragStart(index: number, event: DragEvent): void {
	if (event.dataTransfer) {
		event.dataTransfer.effectAllowed = 'move';
		// Some browsers require data to be set for a drag to start.
		event.dataTransfer.setData('text/plain', String(index));
	}
	controller.dragStart(index);
}

function onDragOver(index: number, event: DragEvent): void {
	// Calling preventDefault marks this element as a valid drop target.
	event.preventDefault();
	if (event.dataTransfer) {
		event.dataTransfer.dropEffect = 'move';
	}
	controller.dragOver(index);
}

function onDrop(index: number, event: DragEvent): void {
	event.preventDefault();
	controller.drop(index, Number.parseInt(event.dataTransfer?.getData('text/plain') ?? '', 10));
}

/**
 * Deliberately does not also `emit('select', index)`: the host's `select` handler navigates
 * the canvas AND closes this whole overlay, which would tear the menu down again on the very
 * right-click that opened it.
 */
function openContextMenu(index: number, event: MouseEvent): void {
	if (controller.openMenu(index, event.clientX, event.clientY)) {
		event.preventDefault();
	}
}

const contextItems = computed<ContextMenuItem[]>(() =>
	view.value.menuEntries.flatMap((entry, position) => {
		const item: ContextMenuItem = {
			id: entry.id,
			label: slideSorterContextMenuLabel(
				t(entry.labelKey),
				entry,
				view.value.selectedIndexes.length,
			),
			disabled: entry.disabled,
		};
		return entry.separatorBefore
			? [{ id: `sep-${position}`, label: '', separator: true }, item]
			: [item];
	}),
);

function onContextSelect(id: string): void {
	controller.run(id as SlideSorterContextMenuCommandId);
}
const onKeyDown = (event: KeyboardEvent): void => {
	controller.handleKeydown(event);
};

onMounted(() => {
	window.addEventListener('keydown', onKeyDown);
});
onBeforeUnmount(() => {
	window.removeEventListener('keydown', onKeyDown);
});
</script>

<template>
	<div class="pptx-vue-sorter" role="dialog" :aria-label="t('pptx.slideSorter.title')">
		<header class="pptx-vue-sorter-head">
			<pptx-ui-slide-clipboard
				:labels.prop="slideClipboardLabels(t)"
				:copyAction.prop="
					copySystemSlides ? () => copySystemSlides!(controller.selectedIndexes()) : undefined
				"
				:pasteAction.prop="
					canEdit && pasteSystemSlides
						? () => pasteSystemSlides!(controller.pasteInsertAt())
						: undefined
				"
			/>
			<h2 class="pptx-vue-sorter-title">{{ t('pptx.slideSorter.title') }}</h2>
			<button
				type="button"
				class="pptx-vue-sorter-close"
				:aria-label="t('pptx.slideSorter.close')"
				@click="emit('close')"
			>
				<X :size="16" aria-hidden="true" />
			</button>
		</header>

		<div
			class="pptx-vue-sorter-grid"
			:style="{ gridTemplateColumns: `repeat(${view.gridColumns}, minmax(0, 1fr))` }"
		>
			<div
				v-for="(slide, index) in slides"
				:key="slide.id ?? index"
				class="pptx-vue-sorter-tile"
				:class="{
					'is-active': view.selectedIds.has(slide.id),
					'is-dragging': index === state.dragIndex,
					'is-drop-target': index === state.dragOverIndex && index !== state.dragIndex,
					'is-hidden': Boolean(slide.hidden),
				}"
				:draggable="canEdit ? 'true' : 'false'"
				data-pptx-chrome="sorter-tile"
				:data-index="index"
				:data-pptx-selected="view.selectedIds.has(slide.id)"
				:data-pptx-slide-hidden="hiddenCue(slide.hidden, 'sorter', index).marker"
				:aria-label="t('pptx.notes.slideN', { n: index + 1 })"
				:aria-current="index === activeIndex ? 'true' : undefined"
				:aria-describedby="hiddenCue(slide.hidden, 'sorter', index).labelId"
				@click="controller.select(index, $event)"
				@dblclick="emit('select', index)"
				@contextmenu="openContextMenu(index, $event)"
				data-pptx-touch-press="menu dblclick"
				@dragstart="onDragStart(index, $event)"
				@dragover="onDragOver(index, $event)"
				@drop="onDrop(index, $event)"
				@dragend="controller.dragEnd()"
			>
				<div class="pptx-vue-sorter-stage" :style="stageWrapStyle" aria-hidden="true">
					<SlideStage
						:slide="slide"
						:canvas-size="canvasSize"
						:media-data-urls="mediaDataUrls"
						:scale="tileScale"
					/>
				</div>
				<span
					class="pptx-vue-sorter-index"
					:style="slide.hidden ? { backgroundImage: slashGradient } : undefined"
					>{{ index + 1 }}</span
				>
				<span
					v-if="slide.hidden"
					:id="hiddenCue(slide.hidden, 'sorter', index).labelId"
					class="pptx-vue-sorter-hidden"
					>{{ t('pptx.slideSorter.hidden') }}</span
				>
			</div>
		</div>

		<label
			>{{ t('pptx.slideSorter.zoom') }}
			<input
				type="range"
				min="50"
				max="200"
				step="10"
				v-model.number="zoom"
				:aria-label="t('pptx.slideSorter.zoom')"
			/>{{ zoom }}%</label
		>
		<ContextMenu
			:open="state.menu !== null"
			:x="state.menu?.x ?? 0"
			:y="state.menu?.y ?? 0"
			:items="contextItems"
			@select="onContextSelect"
			@close="controller.closeMenu()"
		/>
	</div>
</template>

<style scoped src="./slide-sorter.css"></style>
