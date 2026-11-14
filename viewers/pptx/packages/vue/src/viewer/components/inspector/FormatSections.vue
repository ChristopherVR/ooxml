<script setup lang="ts">
/**
 * FormatSections: the Properties tab as a selection-aware Format pane. Draws
 * the shared section list for the selection (`resolveFormatSections`) as
 * `<pptx-ui-format-section>` elements, keeps each one's open state in the
 * viewer's Format pane controller, and reveals a launcher's requested section
 * once rendered. The parent fills each section through the default slot
 * (`{ view, category }`) with its existing panels.
 */
import {
	FORMAT_SECTION_ATTRIBUTE,
	formatSectionOpen,
	isFormatSectionMixed,
	resolveFormatCategory,
	resolveFormatSections,
	revealFormatSection,
} from 'ooxml-ui/pptx';
import type { FormatSectionView, FormatSelectionCategory } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { computed, ref, watch } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectFormatPane } from '../../composables/format-pane-context';
import { useController } from '../../composables/useController';

const props = defineProps<{
	/** The selection, primary first (empty for the slide). */
	elements: readonly PptxElement[];
}>();

defineSlots<{
	default(props: { view: FormatSectionView; category: FormatSelectionCategory }): unknown;
}>();

const { t } = useI18n();
const pane = injectFormatPane();
const state = useController(pane);
const views = computed(() => resolveFormatSections(props.elements));
const category = computed(() => resolveFormatCategory(props.elements));
const root = ref<HTMLElement | null>(null);

function sectionAttrs(view: FormatSectionView): Record<string, unknown> {
	return {
		[FORMAT_SECTION_ATTRIBUTE]: view.id,
		'.label': t(view.labelKey),
		'.open': formatSectionOpen(state.value, category.value, view),
		'.mixed': isFormatSectionMixed(props.elements, view.id),
		'.mixedLabel': t('pptx.formatPane.mixed'),
	};
}

function onToggle(view: FormatSectionView, event: Event): void {
	pane.setOpen(category.value, view.id, (event as CustomEvent<{ open: boolean }>).detail.open);
}

// After each render: take a launcher's request, else scroll to the top on a category change.
watch(
	[category, views, () => state.value.request],
	() => {
		const changed = pane.noteCategory(category.value);
		const reveal = pane.takeRequest(category.value, views.value);
		const node = root.value;
		if (!node) {
			return;
		}
		if (reveal) {
			requestAnimationFrame(() => revealFormatSection(node, reveal.section, reveal));
		} else if (changed) {
			const scroller = node.closest<HTMLElement>(
				'[data-pptx-chrome="inspector-body"], .pptx-vue-inspector',
			);
			if (scroller) {
				scroller.scrollTop = 0;
			}
		}
	},
	{ flush: 'post', immediate: true },
);
</script>

<template>
	<div
		ref="root"
		:data-pptx-format-sections="category"
		role="group"
		:aria-label="t('pptx.formatPane.sections')"
		class="flex flex-col"
	>
		<pptx-ui-format-section
			v-for="view in views"
			:key="view.id"
			v-bind="sectionAttrs(view)"
			@office-section-toggle="onToggle(view, $event)"
		>
			<slot :view="view" :category="category" />
		</pptx-ui-format-section>
	</div>
</template>
