// @vitest-environment happy-dom
/**
 * A drag through one of the inspector's native colour pickers (fill, line,
 * text, or an effects colour field) previews every shade and is ONE undo step
 * and ONE recent colour (shared `createLiveColorInput`), as in the other four
 * bindings.
 *
 * Every shade used to go through the inspector's normal edit path, which arms
 * the undo tracker (`markDirty`), so a drag through twenty shades took twenty
 * presses of Undo. Wired as the viewer wires it: the real `ShapeTextPanels`
 * over the real `useEditorHistory` and `useElementOperations`, with the live
 * edit context `ViewerInspector` provides (`colour-drag.test-support.tsx`).
 */
import type { PptxElement } from 'pptx-viewer-core';
import { act } from 'react';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import type { Control, Styled } from './colour-drag.test-support';
import {
	api,
	colorInput,
	fire as fireInput,
	renderHarness,
	setUpColourDragHarness,
} from './colour-drag.test-support';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string, fallback?: unknown) => (typeof fallback === 'string' ? fallback : key),
	}),
}));

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

setUpColourDragHarness();

beforeEach(() => renderHarness([shape()]));

const COLOR: Record<Control, (el: Styled) => string | undefined> = {
	fill: (el) => el.shapeStyle.fillColor,
	line: (el) => el.shapeStyle.strokeColor,
	text: (el) => el.textStyle?.color,
	glow: (el) => el.shapeStyle.glowColor,
};
const REF: Record<Control, (el: Styled) => unknown> = {
	fill: (el) => el.shapeStyle.fillColorRef,
	line: (el) => el.shapeStyle.strokeColorRef,
	text: (el) => el.textStyle?.colorRef,
	glow: (el) => el.shapeStyle.glowRadius,
};

/** Fire a native colour-input event at the `control` picker. */
function fire(control: Control, type: 'input' | 'change' | 'blur', hex: string): void {
	fireInput(colorInput(control), type, hex);
}

/** Set the value and fire `input` then `change` in one act, so React renders only after both. */
function pick(control: Control, hex: string): void {
	const input = colorInput(control);
	act(() => {
		Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(input, hex);
		input.dispatchEvent(new Event('input', { bubbles: true }));
		input.dispatchEvent(new Event('change', { bubbles: true }));
	});
}

const SHADES = ['#330000', '#660000', '#990000'];

describe.each(['fill', 'line', 'text', 'glow'] as const)('the %s colour picker', (control) => {
	const color = COLOR[control];
	const themeRef = REF[control];

	it('shows a dragged shade on the element at once, with no undo step or recent colour', () => {
		fire(control, 'input', '#330000');

		expect(color(api!.element())).toBe('#330000');
		expect(api!.canUndo()).toBeFalsy();
		expect(api!.recent()).toStrictEqual([]);
	});

	it('makes a drag and its settling change ONE undo step and ONE recent colour', () => {
		const before = api!.element();
		for (const shade of SHADES) {
			fire(control, 'input', shade);
			expect(color(api!.element())).toBe(shade);
		}
		fire(control, 'change', '#990000');

		expect(color(api!.element())).toBe('#990000');
		if (control !== 'glow') {
			// A native pick is a custom colour: the theme ref goes.
			expect(themeRef(api!.element())).toBeUndefined();
		}
		expect(api!.recent()).toStrictEqual(['#990000']);

		act(() => api!.undo());
		expect(color(api!.element())).toBe(color(before));
		expect(themeRef(api!.element())).toStrictEqual(themeRef(before));
		expect(api!.canUndo()).toBeFalsy();
	});

	it('leaves no undo step, and keeps the theme colour, when a drag settles where it started', () => {
		const before = api!.element();
		fire(control, 'input', '#330000');
		fire(control, 'input', color(before)!);
		fire(control, 'change', color(before)!);

		expect(api!.canUndo()).toBeFalsy();
		// The `change` is still a pick: its colour counts as recent, though nothing was edited.
		expect(api!.recent()).toStrictEqual([color(before)]);
		expect(color(api!.element())).toBe(color(before));
		expect(themeRef(api!.element())).toStrictEqual(themeRef(before));
	});

	it('commits a pick whose input and change arrive in one turn as one undo step', () => {
		const before = api!.element();
		pick(control, '#00ff00');

		expect(color(api!.element())).toBe('#00ff00');
		expect(api!.recent()).toStrictEqual(['#00ff00']);
		act(() => api!.undo());
		expect(color(api!.element())).toBe(color(before));
		expect(api!.canUndo()).toBeFalsy();
	});

	it('counts a pick of the colour already there as a recent colour, with no undo step', () => {
		const before = api!.element();
		pick(control, color(before)!);

		expect(api!.recent()).toStrictEqual([color(before)]);
		expect(api!.canUndo()).toBeFalsy();
		expect(themeRef(api!.element())).toStrictEqual(themeRef(before));
	});

	it('settles a drag that no change event ended when the picker loses focus', () => {
		const before = api!.element();
		fire(control, 'input', '#330000');
		fire(control, 'blur', '#330000');

		expect(api!.recent()).toStrictEqual(['#330000']);
		act(() => api!.undo());
		expect(color(api!.element())).toBe(color(before));
		expect(api!.canUndo()).toBeFalsy();
	});
});
