// oxlint-disable react-hooks/rules-of-hooks -- Vue composables, not React hooks
/**
 * A live colour drag in the Vue inspector stays on the element it started on
 * and never overwrites another edit (shared `createLiveColorInput`): the
 * selection moving mid-drag puts the first shape back without landing the pick
 * on the second, an edit to another field of the dragged shape survives the
 * commit and its undo step, an effects colour field tells a new target from a
 * re-render of its own drag, and a drag on a group member or a template
 * element is put back once the selection has left it. Wired as
 * `PowerPointViewer` wires it.
 */
import { disableAutoUnmount, enableAutoUnmount, mount } from '@vue/test-utils';
import type { PptxElement, PptxSlide, ShapeStyle } from 'pptx-viewer-core';
import { afterAll, afterEach, describe, expect, it } from 'vitest';
import type { Component } from 'vue';
import { computed, defineComponent, h, nextTick, provide, ref, shallowRef } from 'vue';

import { InspectorPreviewKey } from '../../composables/inspector-preview-context';
import { RecentColorsKey } from '../../composables/recent-colors-context';
import type { TemplateElementMap } from '../../composables/template-editing';
import { useEditorHistory } from '../../composables/useEditorHistory';
import { useEditorOperations } from '../../composables/useEditorOperations';
import { useInspectorWiring } from '../../composables/useInspectorWiring';
import { useSelectionModel } from '../../composables/useSelectionModel';
import EffectsPanel from './EffectsPanel.vue';
import FillPanel from './FillPanel.vue';

// Every wrapper goes after its test, a failing one too: the worker is shared.
enableAutoUnmount(afterEach);

afterAll(disableAutoUnmount);

const A_STYLE: ShapeStyle = {
	fillMode: 'solid',
	fillColor: '#4472c4',
	fillColorRef: { scheme: 'accent1' },
	strokeColor: '#ed7d31',
	strokeWidth: 2,
	glowColor: '#ff0000',
	glowRadius: 8,
	glowOpacity: 0.5,
};
const B_STYLE: ShapeStyle = {
	fillMode: 'solid',
	fillColor: '#00aa55',
	fillColorRef: { scheme: 'accent6' },
	strokeWidth: 7,
	glowColor: '#ff0000',
	glowRadius: 3,
	glowOpacity: 0.9,
};

function shape(id: string, shapeStyle: ShapeStyle): PptxElement {
	return {
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		shapeStyle,
	} as unknown as PptxElement;
}

function group(id: string, children: PptxElement[]): PptxElement {
	return {
		type: 'group',
		id,
		x: 200,
		y: 100,
		width: 100,
		height: 50,
		children,
	} as unknown as PptxElement;
}

/** The element `id` anywhere in `elements`, group members included. */
function findIn(elements: readonly PptxElement[], id: string): PptxElement | undefined {
	for (const el of elements) {
		const hit = el.id === id ? el : findIn((el as { children?: PptxElement[] }).children ?? [], id);
		if (hit) {
			return hit;
		}
	}
	return undefined;
}

interface Deck {
	/** The slide's elements (shapes `a` and `b` when omitted). */
	elements?: PptxElement[];
	/** The slide's template (layout and master) elements. */
	templates?: PptxElement[];
	/** The element selected first (`a` when omitted). */
	selected?: string;
}

function harness(panel: Component, selector: string, deck: Deck = {}) {
	const slides = shallowRef<PptxSlide[]>([
		{
			id: 's',
			rId: 'rId1',
			slideNumber: 1,
			elements: deck.elements ?? [shape('a', A_STYLE), shape('b', B_STYLE)],
		} as PptxSlide,
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
	const fire = async (type: 'input' | 'change' | 'blur', hex: string): Promise<void> => {
		const input = wrapper.find(selector);
		(input.element as HTMLInputElement).value = hex;
		await input.trigger(type);
		await nextTick();
	};
	const style = (id: string): ShapeStyle =>
		(
			findIn([...slides.value[0]!.elements, ...(templateElementsBySlideId.value.s ?? [])], id) as {
				shapeStyle: ShapeStyle;
			}
		).shapeStyle;
	const select = async (id: string): Promise<void> => {
		selection.selectElement(id, false);
		await nextTick();
	};
	return { slides, history, recent, fire, style, select };
}

describe('a fill drag the selection leaves', () => {
	it('puts the first shape back and does not land the pick on the second', async () => {
		const view = harness(FillPanel, 'input.pptx-vue-fill-color');
		await view.fire('input', '#330000');
		await view.select('b');
		await view.fire('change', '#330000');

		expect(view.style('a')).toStrictEqual(A_STYLE);
		expect(view.style('b')).toStrictEqual(B_STYLE);
		expect(view.history.canUndo.value).toBeFalsy();
		expect(view.recent.value).toStrictEqual([]);
	});
});

describe('a fill drag the selection leaves after another edit to the shape', () => {
	it('puts the colour back on the shape as it is now, keeping that edit', async () => {
		const view = harness(FillPanel, 'input.pptx-vue-fill-color');
		await view.fire('input', '#330000');
		view.slides.value = view.slides.value.map((slide) => ({
			...slide,
			elements: slide.elements.map((el) =>
				el.id === 'a'
					? ({ ...el, shapeStyle: { ...view.style('a'), strokeWidth: 5 } } as PptxElement)
					: el,
			),
		}));
		await nextTick();
		await view.select('b');
		await view.fire('change', '#330000');

		expect(view.style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('another edit to the dragged shape', () => {
	it('survives the commit, and the one undo step keeps it', async () => {
		const view = harness(FillPanel, 'input.pptx-vue-fill-color');
		await view.fire('input', '#330000');
		// A collaborator widens the line while the picker is open.
		view.slides.value = view.slides.value.map((slide) => ({
			...slide,
			elements: slide.elements.map((el) =>
				el.id === 'a'
					? ({ ...el, shapeStyle: { ...view.style('a'), strokeWidth: 5 } } as PptxElement)
					: el,
			),
		}));
		await nextTick();
		await view.fire('change', '#330000');

		expect(view.style('a').strokeWidth).toBe(5);
		expect(view.style('a').fillColor).toBe('#330000');
		view.history.undo();
		await nextTick();
		expect(view.style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('an effects colour field whose target changes mid-drag', () => {
	it('does not write the first target onto the second', async () => {
		const view = harness(EffectsPanel, '[data-testid="fx-glow-color"]');
		await view.fire('input', '#00ff00');
		await view.fire('input', '#ff0000');
		await view.select('b');
		await view.fire('blur', '#ff0000');

		expect(view.style('b')).toStrictEqual(B_STYLE);
		expect(view.style('a')).toMatchObject({
			glowColor: '#ff0000',
			glowRadius: 8,
			glowOpacity: 0.5,
		});
	});

	it('puts the first target back as it is now, keeping an edit made there since', async () => {
		const view = harness(EffectsPanel, '[data-testid="fx-glow-color"]');
		await view.fire('input', '#00ff00');
		view.slides.value = view.slides.value.map((slide) => ({
			...slide,
			elements: slide.elements.map((el) =>
				el.id === 'a'
					? ({ ...el, shapeStyle: { ...view.style('a'), glowRadius: 12 } } as PptxElement)
					: el,
			),
		}));
		await nextTick();
		await view.select('b');
		await view.fire('blur', '#00ff00');

		expect(view.style('a')).toStrictEqual({ ...A_STYLE, glowRadius: 12 });
		expect(view.style('b')).toStrictEqual(B_STYLE);
	});
});

describe('a fill drag the selection leaves on a group member', () => {
	it('puts the member back, with no undo step', async () => {
		const view = harness(FillPanel, 'input.pptx-vue-fill-color', {
			elements: [group('g', [shape('m', A_STYLE)]), shape('b', B_STYLE)],
			selected: 'm',
		});
		await view.fire('input', '#330000');
		expect(view.style('m').fillColor).toBe('#330000');
		await view.select('b');
		await view.fire('change', '#330000');

		expect(view.style('m')).toStrictEqual(A_STYLE);
		expect(view.style('b')).toStrictEqual(B_STYLE);
		expect(view.history.canUndo.value).toBeFalsy();
	});
});

describe('a fill drag the selection leaves on a template element', () => {
	it('puts the template element back, with no undo step', async () => {
		const view = harness(FillPanel, 'input.pptx-vue-fill-color', {
			elements: [shape('b', B_STYLE)],
			templates: [shape('layout-t', A_STYLE)],
			selected: 'layout-t',
		});
		await view.fire('input', '#330000');
		expect(view.style('layout-t').fillColor).toBe('#330000');
		await view.select('b');
		await view.fire('change', '#330000');

		expect(view.style('layout-t')).toStrictEqual(A_STYLE);
		expect(view.style('b')).toStrictEqual(B_STYLE);
		expect(view.history.canUndo.value).toBeFalsy();
	});
});
