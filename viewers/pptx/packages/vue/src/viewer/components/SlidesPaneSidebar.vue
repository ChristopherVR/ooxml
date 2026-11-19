<script setup lang="ts">
import { EyeOff, MessageSquare, Plus } from '@lucide/vue';
import {
	createDeferredRows,
	createVirtualWindow,
	focusCurrentRow,
	railHasFocus,
	HIDDEN_SLIDE_LABEL_KEY,
	HIDDEN_SLIDE_SLASH_GRADIENT,
	hiddenSlideCue,
	SLIDE_VIRTUALIZATION_THRESHOLD,
	editorThumbnailHeight,
	editorThumbnailStep,
	EDITOR_THUMBNAIL_WIDTH,
} from 'ooxml-ui/pptx';
/**
 * SlidesPaneSidebar - Vue port of React's `SlidesPaneSidebar` + `SlideItem`.
 *
 * The flat slide rail: a scrollable list of number-left thumbnails (active slide
 * gets a left primary bar + accent background), drag-to-reorder, a right-click
 * context menu (Duplicate / Delete / Hide), and a bottom "Add slide" button.
 * Class strings are copied from React's `SlideItem`/`SlidesPaneSidebar` for
 * visual parity; `react-icons/lu` glyphs map to `@lucide/vue`.
 *
 * Sectioned decks keep using `SectionList`; this renders the non-sectioned case
 * and virtualizes decks at the same 50-slide threshold as React.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { computed, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue';
import { useI18n } from 'vue-i18n';

import { cn } from '../../utils';
import { useSlidePaneRailMenu } from '../composables/useSlidePaneRailMenu';
import type { CanvasSize } from '../types';
import ContextMenu from './ContextMenu.vue';
import SlideStage from './SlideStage.vue';

const props = withDefaults(
	defineProps<{
		slides: PptxSlide[];
		activeIndex: number;
		canvasSize: CanvasSize;
		mediaDataUrls: Map<string, string>;
		canEdit: boolean;
		/** Thumbnail content width in px (matches React's SLIDE_NAV_THUMBNAIL_WIDTH). */
		thumbWidth?: number;
		/** Fill the container instead of the rail width (the phone Slides sheet). */
		fullWidth?: boolean;
	}>(),
	{ thumbWidth: EDITOR_THUMBNAIL_WIDTH },
);

const emit = defineEmits<{
	select: [index: number];
	reorder: [payload: { from: number; to: number }];
	'add-slide': [];
	'add-slide-after': [index: number];
	duplicate: [indexes: number[]];
	delete: [indexes: number[]];
	'toggle-hidden': [indexes: number[]];
	layout: [payload: { index: number; x: number; y: number }];
	'add-section': [index: number];
}>();

const { t } = useI18n();

/**
 * The shared rail cue for a hidden slide. Exposed to the template rather than
 * inlined so the dim / slash / description all come off one source and cannot
 * drift from the other four bindings.
 */
const hiddenCue = hiddenSlideCue;
const hiddenLabelKey = HIDDEN_SLIDE_LABEL_KEY;
const slashGradient = HIDDEN_SLIDE_SLASH_GRADIENT;

const scale = computed(() => props.thumbWidth / Math.max(1, props.canvasSize.width));
const previewHeight = computed(() =>
	editorThumbnailHeight(props.canvasSize.width, props.canvasSize.height, props.thumbWidth),
);
const listEl = ref<HTMLElement | null>(null);
const itemHeight = computed(() =>
	editorThumbnailStep(props.canvasSize.width, props.canvasSize.height, props.thumbWidth),
);
const shouldVirtualize = computed(() => props.slides.length >= SLIDE_VIRTUALIZATION_THRESHOLD);
// The shared window reads scroll once per frame and reports only a change of
// rendered rows, so the rail re-renders when a row enters, not per scroll event.
const virtualWindow = createVirtualWindow(
	{ totalItems: props.slides.length, itemHeight: itemHeight.value },
	(range) => {
		virtualRange.value = range;
	},
);
const virtualRange = shallowRef(virtualWindow.range);
watch([() => props.slides.length, itemHeight], ([totalItems, height]) =>
	virtualWindow.setShape({ totalItems, itemHeight: height }),
);
watch(listEl, (el) => virtualWindow.attach(el), { flush: 'sync' });
// Rows new to the window build their thumbnail once scrolling pauses.
const readyRows = shallowRef<ReadonlySet<string>>(new Set());
const deferredRows = createDeferredRows(virtualWindow, (ready) => {
	readyRows.value = ready;
});
onBeforeUnmount(() => {
	deferredRows.destroy();
	virtualWindow.destroy();
});
const renderedSlides = computed(() => {
	const start = shouldVirtualize.value ? virtualRange.value.startIndex : 0;
	const end = shouldVirtualize.value ? virtualRange.value.endIndex : props.slides.length - 1;
	return props.slides
		.slice(start, end + 1)
		.map((slide, offset) => ({ slide, index: start + offset }));
});
watch(
	renderedSlides,
	(rows) => deferredRows.sync(shouldVirtualize.value ? rows.map(({ slide }) => slide.id) : []),
	{ immediate: true },
);
watch(
	[() => props.activeIndex, itemHeight],
	async ([index]) => {
		await nextTick();
		// Focus stays on the current row, as in PowerPoint's thumbnail pane: read
		// before the reveal can drop the focused row, moved once the row renders.
		const keepFocus = railHasFocus(listEl.value);
		if (shouldVirtualize.value) {
			virtualWindow.reveal(index);
		}
		if (keepFocus) {
			await nextTick();
			focusCurrentRow(listEl.value);
		}
	},
	{ immediate: true },
);
// A change of slides can remove the focused row (an undone New Slide): focus moves to the
// current row rather than falling to the page, where the editor's shortcuts no longer reach.
watch(
	() => props.slides,
	async () => {
		await nextTick();
		if (railHasFocus(listEl.value)) {
			focusCurrentRow(listEl.value);
		}
	},
);

// ── Drag-to-reorder ──
const dragFrom = ref<number | null>(null);
function onDragStart(e: DragEvent, index: number): void {
	dragFrom.value = index;
	if (e.dataTransfer) {
		e.dataTransfer.effectAllowed = 'move';
	}
}
function onDrop(index: number): void {
	if (dragFrom.value !== null && dragFrom.value !== index) {
		emit('reorder', { from: dragFrom.value, to: index });
	}
	dragFrom.value = null;
}

// ── Multi-select + thumbnail context menu ──
const {
	selectedIds,
	isSelected,
	onSlideClick,
	onPaneKeydown,
	menu,
	closeMenu,
	menuItems,
	onContextMenu,
	onMenuSelect,
} = useSlidePaneRailMenu(
	() => props.slides,
	() => props.activeIndex,
	() => props.canEdit,
	emit,
);
</script>

<template>
	<aside
		role="navigation"
		:aria-label="t('pptx.sections.slides')"
		class="flex h-full flex-col border-r border-border bg-secondary/30 shrink-0"
		:style="{ width: fullWidth ? '100%' : `${thumbWidth + 48}px` }"
		data-pptx-chrome="slides"
		@keydown="onPaneKeydown"
	>
		<div
			data-pptx-chrome="slide-list"
			ref="listEl"
			class="flex-1 overflow-y-auto px-1.5 pb-2 pt-1.5"
		>
			<div
				:data-virtualized="shouldVirtualize ? 'true' : undefined"
				:style="
					shouldVirtualize
						? { height: `${virtualRange.totalHeight}px`, position: 'relative' }
						: undefined
				"
			>
				<div
					class="space-y-1"
					data-pptx-chrome="slide-window"
					:role="shouldVirtualize ? 'list' : undefined"
					:style="
						shouldVirtualize
							? {
									position: 'absolute',
									insetInline: '0',
									top: `${virtualRange.offsetY}px`,
								}
							: undefined
					"
				>
					<!-- The window holds only some rows, so each says where it sits in the deck. -->
					<div
						v-for="{ slide, index } in renderedSlides"
						:key="slide.id ?? index"
						:role="shouldVirtualize ? 'listitem' : undefined"
						:aria-setsize="shouldVirtualize ? slides.length : undefined"
						:aria-posinset="shouldVirtualize ? index + 1 : undefined"
					>
						<button
							type="button"
							:class="
								cn(
									'group relative flex w-full items-center gap-1 cursor-pointer border-0 bg-transparent py-0.5 px-1 text-left transition-all',
									index === activeIndex &&
										'bg-accent/40 before:absolute before:left-0 before:top-1 before:bottom-1 before:w-[3px] before:bg-primary before:rounded-r',
									isSelected(slide.id) &&
										index !== activeIndex &&
										'bg-accent/20 ring-1 ring-inset ring-primary/50',
									slide.hidden && 'opacity-50',
								)
							"
							:draggable="canEdit"
							:data-slide-index="index"
							:aria-label="t('pptx.slidesPanel.goToSlide', { n: index + 1 })"
							data-pptx-chrome="slide-row"
							:aria-current="index === activeIndex ? 'true' : undefined"
							:aria-describedby="hiddenCue(slide.hidden, 'rail', index).labelId"
							:data-pptx-slide-hidden="hiddenCue(slide.hidden, 'rail', index).marker"
							@click="onSlideClick($event, index)"
							@contextmenu="onContextMenu($event, index)"
							@dragstart="onDragStart($event, index)"
							@dragover.prevent
							@drop="onDrop(index)"
						>
							<div
								data-pptx-chrome="slide-number"
								class="flex flex-col items-center gap-0.5 w-5 shrink-0"
							>
								<span
									:class="
										cn(
											'text-[10px] tabular-nums text-right select-none w-full',
											index === activeIndex ? 'text-primary font-medium' : 'text-muted-foreground',
										)
									"
									:style="slide.hidden ? { backgroundImage: slashGradient } : undefined"
								>
									{{ index + 1 }}
								</span>
							</div>

							<!-- Thumbnail -->
							<div
								:class="
									cn(
										'relative flex-1 overflow-hidden border transition-colors bg-white',
										index === activeIndex
											? 'border-primary/60'
											: 'border-transparent group-hover:border-border/40',
									)
								"
								data-pptx-chrome="slide-frame"
								:style="{ width: `${thumbWidth}px`, height: `${previewHeight}px`, flex: 'none' }"
							>
								<SlideStage
									v-if="!shouldVirtualize || readyRows.has(slide.id)"
									:slide="slide"
									:canvas-size="canvasSize"
									:media-data-urls="mediaDataUrls"
									:scale="scale"
								/>
								<div
									v-if="(slide.comments?.length ?? 0) > 0"
									class="absolute top-0.5 right-0.5 flex items-center gap-0.5 rounded bg-amber-500/90 px-1 py-0.5 text-[8px] font-medium text-white leading-none z-10"
								>
									<MessageSquare class="w-2 h-2" />
									{{ slide.comments?.length }}
								</div>
								<div
									v-if="slide.hidden"
									:id="hiddenCue(slide.hidden, 'rail', index).labelId"
									class="absolute bottom-0.5 right-0.5 z-10"
								>
									<EyeOff class="w-3 h-3 text-muted-foreground" />
									<span class="sr-only">{{ t(hiddenLabelKey) }}</span>
								</div>
							</div>
						</button>
					</div>
				</div>
			</div>
		</div>

		<!-- Bottom: Add Slide -->
		<div
			v-if="canEdit"
			data-pptx-chrome="slide-footer"
			class="border-t border-border/60 px-2 py-1.5"
		>
			<button
				type="button"
				class="flex w-full items-center justify-center gap-1 rounded-sm px-2 py-1 text-[11px] text-muted-foreground hover:bg-accent hover:text-foreground transition-colors"
				@click="emit('add-slide')"
			>
				<Plus class="h-3 w-3" />
				{{ t('pptx.sections.addSlide') }}
			</button>
		</div>

		<ContextMenu
			:open="menu.open"
			:x="menu.x"
			:y="menu.y"
			:items="menuItems"
			:aria-label="t('pptx.slidesPane.contextMenu.newSlide')"
			:markers="['data-pptx-slide-pane-context-menu']"
			@select="onMenuSelect"
			@close="closeMenu()"
		/>
	</aside>
</template>
