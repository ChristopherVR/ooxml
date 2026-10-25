/**
 * A drag through one of the inspector's native colour pickers (fill, line,
 * text, or an effects colour field) previews every shade and is ONE undo step
 * and ONE recent colour (shared `createLiveColorInput`). Each `input` used to
 * emit an `update` that the viewer recorded, so a drag through twenty shades
 * took twenty presses of Undo (and the settling `change` one more). Wired as
 * `PowerPointViewer` wires it (`colour-drag.test-support.ts`).
 */
import { disableAutoUnmount, enableAutoUnmount } from '@vue/test-utils';
import type { PptxElement, ShapeStyle, TextStyle } from 'pptx-viewer-core';
import { afterAll, afterEach, describe, expect, it } from 'vitest';
import type { Component } from 'vue';
import { nextTick } from 'vue';

import { mountColourPanel } from './colour-drag.test-support';
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
	const { panel, selector } = PANELS[control];
	const view = mountColourPanel(panel, selector, { elements: [shape()] });
	/** Set the value and fire `input` then `change` in one turn, with no re-render between. */
	const pick = async (hex: string): Promise<void> => {
		const input = view.input();
		(input.element as HTMLInputElement).value = hex;
		input.element.dispatchEvent(new Event('input', { bubbles: true }));
		input.element.dispatchEvent(new Event('change', { bubbles: true }));
		await nextTick();
	};
	return {
		history: view.history,
		fire: view.fire,
		pick,
		recent: view.recent,
		current: () => view.slides.value[0]!.elements[0] as Styled,
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
