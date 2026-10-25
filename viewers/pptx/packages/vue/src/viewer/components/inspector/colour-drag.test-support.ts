// oxlint-disable react-hooks/rules-of-hooks -- Vue composables, not React hooks
/**
 * colour-drag.test-support.ts: an inspector panel mounted for the Vue live
 * colour drag suites, wired as `PowerPointViewer` wires it: the real selection
 * model, history, editor operations and inspector wiring, with the preview
 * channel and recent colours provided.
 *
 * Not a `.test.ts` file on purpose: a shared harness is not a suite.
 */
import { mount } from '@vue/test-utils';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import type { Component } from 'vue';
import { computed, defineComponent, h, nextTick, provide, ref, shallowRef } from 'vue';

import { InspectorPreviewKey } from '../../composables/inspector-preview-context';
import { RecentColorsKey } from '../../composables/recent-colors-context';
import type { TemplateElementMap } from '../../composables/template-editing';
import { useEditorHistory } from '../../composables/useEditorHistory';
import { useEditorOperations } from '../../composables/useEditorOperations';
import { useInspectorWiring } from '../../composables/useInspectorWiring';
import { useSelectionModel } from '../../composables/useSelectionModel';

export interface Deck {
	/** The slide's elements. */
	elements: PptxElement[];
	/** The slide's template (layout and master) elements. */
	templates?: PptxElement[];
	/** The element selected first (`a` when omitted). */
	selected?: string;
}

/** Mount `panel` for the selection of a one-slide `deck`; `selector` finds its colour input. */
export function mountColourPanel(panel: Component, selector: string, deck: Deck) {
	const slides = shallowRef<PptxSlide[]>([
		{ id: 's', rId: 'rId1', slideNumber: 1, elements: deck.elements } as PptxSlide,
	]);
	const templateElementsBySlideId = shallowRef<TemplateElementMap>({ s: deck.templates ?? [] });
	const activeSlideIndex = ref(0);
	const activeSlide = computed(() => slides.value[activeSlideIndex.value]);
	const selection = useSelectionModel({ slides, templateElementsBySlideId, activeSlide });
	selection.selectElement(deck.selected ?? 'a', false);
	const recent = ref<string[]>([]);
	const history = useEditorHistory(slides, templateElementsBySlideId);
	const Host = defineComponent({
		setup() {
			const ops = useEditorOperations({
				slides,
				activeSlideIndex,
				pushHistory: history.pushHistory,
				selectedElementIds: selection.selectedElementIds,
				templateElementsBySlideId,
			});
			const inspector = useInspectorWiring({
				slides,
				activeSlide,
				activeSlideIndex,
				selectedElements: selection.selectedElements,
				pushHistory: history.pushHistory,
				updateElement: (id, patch) => ops.updateElement(id, patch),
				previewElement: (id, patch) => ops.previewElement(id, patch),
				findActiveElement: selection.findActiveElement,
			});
			provide(InspectorPreviewKey, inspector.inspectorPreview);
			provide(RecentColorsKey, {
				recent,
				push: (hex: string) => {
					recent.value = [hex, ...recent.value.filter((c) => c !== hex)];
				},
			} as never);
			return () =>
				h(panel, {
					element: inspector.inspectorElementForPanels.value,
					onUpdate: inspector.onInspectorUpdate,
				});
		},
	});
	const wrapper = mount(Host);
	/** The panel's colour input as rendered now. */
	const input = () => wrapper.find(selector);
	const fire = async (type: 'input' | 'change' | 'blur', hex: string): Promise<void> => {
		const target = input();
		(target.element as HTMLInputElement).value = hex;
		await target.trigger(type);
		await nextTick();
	};
	const select = async (id: string): Promise<void> => {
		selection.selectElement(id, false);
		await nextTick();
	};
	return { slides, templateElementsBySlideId, history, recent, input, fire, select };
}
