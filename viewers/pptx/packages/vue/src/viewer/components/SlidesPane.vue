<script setup lang="ts">
/**
 * SlidesPane: the slide thumbnail rail as a standalone building block.
 *
 * The same rail `PowerPointViewer` shows beside its canvas (it renders this
 * component through `ViewerSlideRail`), driven entirely by flat props and
 * events so a host can place it beside its own `SlideCanvas`. Two shapes,
 * chosen by whether any section group names a section: a flat number-left
 * thumbnail list, or the same thumbnails grouped under collapsible section
 * headers. Thumbnails render the slides with their template (master/layout)
 * layer merged in, and resolve fields and table styles from the deck props,
 * so they draw what the canvas draws.
 *
 * The rail never mutates the deck itself: every command is an event the host
 * maps onto its own slide and section operations. Multi-selection commands
 * arrive as one event per slide, highest index first, so each duplicate lands
 * right after its source and deletions never shift the indexes still to go.
 */
import { Plus } from 'lucide-vue-next';
import {
	buildSaveSlides,
	createSlidesPaneBulkSlideOps,
	EDITOR_THUMBNAIL_WIDTH,
	openSlidesPaneLayoutGallery,
} from 'ooxml-ui/pptx';
import type { SlidesPaneLayoutGalleryRequest } from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxElement,
	PptxHeaderFooter,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { provideSlidesPaneContexts } from '../composables/slides-pane-contexts';
import type { SectionGroup } from '../composables/useSectionOperations';
import type { CanvasSize } from '../types';
import SectionList from './SectionList.vue';
import SlidesPaneSidebar from './SlidesPaneSidebar.vue';

const props = withDefaults(
	defineProps<{
		/** The deck's slides, in order (as `useLoadContent().slides` holds them). */
		slides: PptxSlide[];
		/**
		 * Master/layout elements per slide id (`useLoadContent().templateElementsBySlideId`),
		 * merged behind each thumbnail's own elements as the full viewer renders them.
		 * Omit it when `slides` already carry their template layer.
		 */
		templateElementsBySlideId?: Record<string, PptxElement[]>;
		/** Zero-based index of the active slide; its row is highlighted and `aria-current`. */
		activeSlideIndex: number;
		/** Slide size in px, for the thumbnail aspect ratio and scale. */
		canvasSize: CanvasSize;
		/** Media data URLs by part path (`useLoadContent().mediaDataUrls`), for pictures and posters. */
		mediaDataUrls?: Map<string, string>;
		/**
		 * Slides grouped by section (`useSectionOperations().slidesBySection`). The rail shows
		 * section headers when any group names a section, and the flat list otherwise.
		 */
		sectionGroups?: SectionGroup[];
		/** Show the edit affordances: the "New slide" footer, drag-to-reorder and the context menus. */
		canEdit?: boolean;
		/** Deck header/footer settings, for the thumbnails' date, footer and header fields. */
		headerFooter?: PptxHeaderFooter;
		/** Custom document properties, for the thumbnails' `docproperty` fields. */
		customProperties?: readonly PptxCustomProperty[];
		/** `p:presentation/@firstSlideNum`, so slide-number fields count from it; 1 when omitted. */
		firstSlideNumber?: number;
		/** Locale for the thumbnails' date and time fields (browser default when omitted). */
		locale?: string;
		/** Presentation theme, for table colours, fonts and borders in thumbnails. */
		theme?: PptxTheme;
		/** Parsed table styles (`useLoadContent().tableStyleMap`), for tables in thumbnails. */
		tableStyleMap?: ParsedTableStyleMap;
	}>(),
	{
		templateElementsBySlideId: () => ({}),
		mediaDataUrls: () => new Map<string, string>(),
		sectionGroups: () => [],
		canEdit: false,
		headerFooter: undefined,
		customProperties: undefined,
		firstSlideNumber: undefined,
		locale: undefined,
		theme: undefined,
		tableStyleMap: undefined,
	},
);

const emit = defineEmits<{
	/** Make `index` the active slide (a thumbnail click, or the Layout entry). */
	select: [index: number];
	/** Move the slide at `from` to position `to` (drag-to-reorder). */
	'move-slide': [from: number, to: number];
	/** Add a slide after the active one (the "New slide" footer). */
	'add-slide': [];
	/** Add a slide after `index` (Enter on a thumbnail, or "New Slide" in its menu). */
	'add-slide-after': [index: number];
	/** Duplicate the slide at `index`; once per selected slide, highest index first. */
	'duplicate-slide': [index: number];
	/** Delete the slide at `index`; once per selected slide, highest index first. */
	'delete-slide': [index: number];
	/** Toggle the hidden flag of the slide at `index`; once per selected slide. */
	'toggle-slide-hidden': [index: number];
	/**
	 * The context menu's "Layout" entry, emitted after `select` made the slide active.
	 * Open a Layout gallery anchored at the click point.
	 */
	'open-layout-gallery': [request: SlidesPaneLayoutGalleryRequest];
	/** Start a section named `name` at the slide after `afterSlideIndex`. */
	'add-section': [name: string, afterSlideIndex: number];
	/** Rename a section. */
	'rename-section': [sectionId: string, name: string];
	/** Delete a section (its slides stay in the deck). */
	'delete-section': [sectionId: string];
	/** Move a section one position earlier. */
	'move-section-up': [sectionId: string];
	/** Move a section one position later. */
	'move-section-down': [sectionId: string];
	/** Collapse or expand a section's thumbnails. */
	'toggle-section-collapse': [sectionId: string];
}>();

const { t } = useI18n();

provideSlidesPaneContexts(() => props);

const THUMB_WIDTH = EDITOR_THUMBNAIL_WIDTH;

const mergedSlides = computed(() => buildSaveSlides(props.slides, props.templateElementsBySlideId));
const mergedSlideById = computed(() => new Map(mergedSlides.value.map((s) => [s.id, s])));
const hasSections = computed(() => props.sectionGroups.some((group) => group.section));
// Grouping and order come from the host; only the slide objects are swapped
// for their merged equivalents.
const mergedGroups = computed(() =>
	props.sectionGroups.map((group) => ({
		...group,
		slides: group.slides.map((slide) => mergedSlideById.value.get(slide.id) ?? slide),
	})),
);

const bulk = createSlidesPaneBulkSlideOps({
	duplicateSlide: (index) => emit('duplicate-slide', index),
	deleteSlide: (index) => emit('delete-slide', index),
	toggleSlideHidden: (index) => emit('toggle-slide-hidden', index),
});

function openLayout(payload: { index: number; x: number; y: number }): void {
	openSlidesPaneLayoutGallery(
		{ slideIndex: payload.index, x: payload.x, y: payload.y },
		(index) => emit('select', index),
		(request) => emit('open-layout-gallery', request),
	);
}
function addSection(afterSlideIndex: number): void {
	emit('add-section', t('pptx.sections.defaultName'), afterSlideIndex);
}
</script>

<template>
	<SlidesPaneSidebar
		v-if="!hasSections"
		:slides="mergedSlides"
		:active-index="activeSlideIndex"
		:canvas-size="canvasSize"
		:media-data-urls="mediaDataUrls"
		:can-edit="canEdit"
		:thumb-width="THUMB_WIDTH"
		@select="(i) => emit('select', i)"
		@reorder="(p) => emit('move-slide', p.from, p.to)"
		@add-slide="emit('add-slide')"
		@add-slide-after="(i) => emit('add-slide-after', i)"
		@duplicate="bulk.duplicateSlides"
		@delete="bulk.deleteSlides"
		@toggle-hidden="bulk.toggleHideSlides"
		@layout="openLayout"
		@add-section="addSection"
	/>
	<nav
		v-else
		class="pptx-vue-thumbnails"
		data-pptx-chrome="slides"
		:aria-label="t('pptx.sections.slides')"
	>
		<SectionList
			:groups="mergedGroups"
			:canvas-size="canvasSize"
			:media-data-urls="mediaDataUrls"
			:active-index="activeSlideIndex"
			:can-edit="canEdit"
			@select="(i) => emit('select', i)"
			@toggle-collapse="(id) => emit('toggle-section-collapse', id)"
			@rename="(id, name) => emit('rename-section', id, name)"
			@move-up="(id) => emit('move-section-up', id)"
			@move-down="(id) => emit('move-section-down', id)"
			@delete-section="(id) => emit('delete-section', id)"
			@add-section="addSection"
			@add-slide-after="(i) => emit('add-slide-after', i)"
			@duplicate="bulk.duplicateSlides"
			@delete="bulk.deleteSlides"
			@toggle-hidden="bulk.toggleHideSlides"
			@layout="openLayout"
		/>
		<!-- The pinned Add Slide footer, the one persistent rail action in every binding. -->
		<div
			v-if="canEdit"
			data-pptx-chrome="slide-footer"
			class="sticky bottom-0 border-t border-border/60 bg-card px-2 py-1.5"
		>
			<button
				type="button"
				class="flex w-full items-center justify-center gap-1 rounded-sm px-2 py-1 text-[11px] text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
				@click="emit('add-slide')"
			>
				<Plus class="h-3 w-3" />
				{{ t('pptx.sections.addSlide') }}
			</button>
		</div>
	</nav>
</template>
