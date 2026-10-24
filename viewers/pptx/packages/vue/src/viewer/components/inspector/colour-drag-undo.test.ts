// oxlint-disable react-hooks/rules-of-hooks -- Vue composables, not React hooks
/**
 * A drag through one of the inspector's native colour pickers (fill, line,
 * text, or an effects colour field) previews every shade and is ONE undo step
 * and ONE recent colour (shared `createLiveColorInput`). Each `input` used to
 * emit an `update` that the viewer recorded, so a drag through twenty shades
 * took twenty presses of Undo (and the settling `change` one more). Wired as
 * `PowerPointViewer` wires it: the real history, editor operations and
 * inspector wiring, with the preview channel and recent colours provided.
 */
import { disableAutoUnmount, enableAutoUnmount, mount } from '@vue/test-utils';
import type { PptxElement, PptxSlide, ShapeStyle, TextStyle } from 'pptx-viewer-core';
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
import StrokePanel from './StrokePanel.vue';
import TextPanel from './TextPanel.vue';

// Every wrapper goes after its test, a failing one too: the worker is shared.
enableAutoUnmount(afterEach);

afterAll(disableAutoUnmount);

type Styled = PptxElement & { shapeStyle: ShapeStyle; textStyle: TextStyle };

function shape(): PptxElement {
	return {
		type: 'shape',
		id: 'a',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		text: 'Hi',
		shapeStyle: {
			fillMode: 'solid',
			fillColor: '#4472c4',
			fillColorRef: { scheme: 'accent1' },
			strokeColor: '#ed7d31',
			strokeColorRef: { scheme: 'accent2' },
			strokeWidth: 2,
			glowColor: '#ff0000',
			glowRadius: 8,
			glowOpacity: 0.5,
		},
		textStyle: { color: '#112233', colorRef: { scheme: 'tx1' } },
	} as unknown as PptxElement;
}

type Control = 'fill' | 'line' | 'text' | 'glow';

const PANELS: Record<Control, { panel: Component; selector: string }> = {
	fill: { panel: FillPanel, selector: 'input.pptx-vue-fill-color' },
	line: { panel: StrokePanel, selector: 'input.pptx-vue-stroke-color' },
	text: { panel: TextPanel, selector: 'input.pptx-vue-text-color' },
	glow: { panel: EffectsPanel, selector: '[data-testid="fx-glow-color"]' },
};

function harness(control: Control) {
	const slides = shallowRef<PptxSlide[]>([
		{ id: 's', rId: 'rId1', slideNumber: 1, elements: [shape()] } as PptxSlide,
	]);
	const templateElementsBySlideId = shallowRef<TemplateElementMap>({});
	const activeSlideIndex = ref(0);
	const activeSlide = computed(() => slides.value[activeSlideIndex.value]);
	const selection = useSelectionModel({ slides, templateElementsBySlideId, activeSlide });
	selection.selectElement('a', false);
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
				h(PANELS[control].panel, {
					element: inspector.inspectorElementForPanels.value,
					onUpdate: inspector.onInspectorUpdate,
				});
		},
	});
	const wrapper = mount(Host);
	const fire = async (type: 'input' | 'change' | 'blur', hex: string): Promise<void> => {
		const input = wrapper.find(PANELS[control].selector);
		(input.element as HTMLInputElement).value = hex;
		await input.trigger(type);
		await nextTick();
	};
	/** Set the value and fire `input` then `change` in one turn, with no re-render between. */
	const pick = async (hex: string): Promise<void> => {
		const input = wrapper.find(PANELS[control].selector);
		(input.element as HTMLInputElement).value = hex;
		input.element.dispatchEvent(new Event('input', { bubbles: true }));
		input.element.dispatchEvent(new Event('change', { bubbles: true }));
		await nextTick();
	};
	return {
		history,
		fire,
		pick,
		recent,
		current: () => slides.value[0]!.elements[0] as Styled,
	};
}

const COLOR: Record<Control, (el: Styled) => string | undefined> = {
	fill: (el) => el.shapeStyle.fillColor,
	line: (el) => el.shapeStyle.strokeColor,
	text: (el) => el.textStyle.color,
	glow: (el) => el.shapeStyle.glowColor,
};
const REF: Record<Control, (el: Styled) => unknown> = {
	fill: (el) => el.shapeStyle.fillColorRef,
	line: (el) => el.shapeStyle.strokeColorRef,
	text: (el) => el.textStyle.colorRef,
	glow: (el) => el.shapeStyle.glowRadius,
};

const SHADES = ['#330000', '#660000', '#990000'];

describe.each(['fill', 'line', 'text', 'glow'] as const)('the %s colour picker', (control) => {
	const color = COLOR[control];
	const themeRef = REF[control];

	it('shows a dragged shade on the element at once, with no undo step or recent colour', async () => {
		const view = harness(control);
		await view.fire('input', '#330000');

		expect(color(view.current())).toBe('#330000');
		expect(view.history.canUndo.value).toBeFalsy();
		expect(view.recent.value).toStrictEqual([]);
	});

	it('makes a drag and its settling change ONE undo step and ONE recent colour', async () => {
		const view = harness(control);
		const before = view.current();
		for (const shade of SHADES) {
			await view.fire('input', shade);
			expect(color(view.current())).toBe(shade);
		}
		await view.fire('change', '#990000');

		expect(color(view.current())).toBe('#990000');
		if (control !== 'glow') {
			// A native pick is a custom colour: the theme ref goes.
			expect(themeRef(view.current())).toBeUndefined();
		}
		expect(view.recent.value).toStrictEqual(['#990000']);

		view.history.undo();
		await nextTick();
		expect(color(view.current())).toBe(color(before));
		expect(themeRef(view.current())).toStrictEqual(themeRef(before));
		expect(view.history.canUndo.value).toBeFalsy();
	});

	it('leaves no undo step, and keeps the theme colour, when a drag settles where it started', async () => {
		const view = harness(control);
		const before = view.current();
		await view.fire('input', '#330000');
		await view.fire('input', color(before)!);
		await view.fire('change', color(before)!);

		expect(view.history.canUndo.value).toBeFalsy();
		// The `change` is still a pick: its colour counts as recent, though nothing was edited.
		expect(view.recent.value).toStrictEqual([color(before)]);
		expect(color(view.current())).toBe(color(before));
		expect(themeRef(view.current())).toStrictEqual(themeRef(before));
	});

	it('commits a pick whose input and change arrive in one turn as one undo step', async () => {
		const view = harness(control);
		const before = view.current();
		await view.pick('#00ff00');

		expect(color(view.current())).toBe('#00ff00');
		expect(view.recent.value).toStrictEqual(['#00ff00']);
		view.history.undo();
		await nextTick();
		expect(color(view.current())).toBe(color(before));
		expect(view.history.canUndo.value).toBeFalsy();
	});

	it('counts a pick of the colour already there as a recent colour, with no undo step', async () => {
		const view = harness(control);
		const before = view.current();
		await view.pick(color(before)!);

		expect(view.recent.value).toStrictEqual([color(before)]);
		expect(view.history.canUndo.value).toBeFalsy();
		expect(themeRef(view.current())).toStrictEqual(themeRef(before));
	});

	it('settles a drag that no change event ended when the picker loses focus', async () => {
		const view = harness(control);
		const before = view.current();
		await view.fire('input', '#330000');
		await view.fire('blur', '#330000');

		expect(view.recent.value).toStrictEqual(['#330000']);
		view.history.undo();
		await nextTick();
		expect(color(view.current())).toBe(color(before));
		expect(view.history.canUndo.value).toBeFalsy();
	});
});
