/**
 * A drag through one of the inspector's native colour pickers (fill, line,
 * text, or an effects colour field) previews every shade and is ONE undo step
 * and ONE recent colour (shared `createLiveColorInput`). Each `input` used to go
 * through the recorded edit path, so a drag through twenty shades took twenty
 * presses of Undo to take back. Wired end to end: the sections, the chrome's
 * inspector handlers and the real edit actions and history over one store.
 */
import type { PptxElement, PptxSlide, ShapeStyle, TextStyle } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { buildChromeCallbacks } from '../../chrome-callbacks';
import type { ChromeCallbackDeps } from '../../chrome-callbacks';
import { createEditActions } from '../../editor/editor-edit-ops';
import { createEditorOps } from '../../editor/editor-operations';
import { createTranslator } from '../../i18n';
import { createInitialViewerState, createStore } from '../../state';
import { createFillSection } from './fill-section';
import { createTextSection } from './text-section';
import type { InspectorState } from './types';

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

function mount() {
	const slide = {
		id: 's',
		rId: 'rId1',
		slideNumber: 1,
		notes: '',
		elements: [shape()],
	} as PptxSlide;
	const store = createStore({
		...createInitialViewerState(),
		slides: [slide],
		currentSlide: 0,
		editable: true,
		selectedElementId: 'a',
		presentationProperties: { mruColors: [] },
	});
	const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
	const actions = createEditActions({
		doc: document,
		getTranslator: () => createTranslator(),
		store,
		ops,
		getHandler: () => null,
		setHandler: () => {},
	});
	const deps = new Proxy({ getEditActions: () => actions } as unknown as ChromeCallbackDeps, {
		get: (target, property, receiver) => Reflect.get(target, property, receiver) ?? vi.fn(),
	});
	const handlers = buildChromeCallbacks(deps).inspectorHandlers;
	const section = () => document.createElement('section');
	const fill = createFillSection(document, (key) => key, section, handlers);
	const text = createTextSection(document, (key) => key, section, handlers);
	const current = () => store.get().slides[0]!.elements[0] as Styled;
	/** Repaint both sections from the store, as the viewer does after each change. */
	const render = (): void => {
		const el = current();
		const state = {
			hasSelection: true,
			elementId: el.id,
			canShape: true,
			canText: true,
			fillColor: el.shapeStyle.fillColor,
			strokeColor: el.shapeStyle.strokeColor,
			shapeStyle: el.shapeStyle,
			textStyle: el.textStyle,
			fillOpacity: 1,
			gradientEnabled: false,
			gradient: { type: 'linear', angle: 90, stops: [] },
			recentColors: [],
		} as unknown as InspectorState;
		fill.update(state);
		text.update(state);
	};
	render();
	const [fillInput, lineInput] = Array.from(
		fill.el.querySelectorAll<HTMLInputElement>(':scope > [role="group"] input[type="color"]'),
	);
	const inputs: Record<Control, HTMLInputElement | null> = {
		fill: fillInput ?? null,
		line: lineInput ?? null,
		text: text.el.querySelector<HTMLInputElement>('input[type="color"]'),
		glow: fill.el.querySelector<HTMLInputElement>('[data-testid="fx-glow-color"]'),
	};
	const fire = (control: Control, type: 'input' | 'change' | 'blur', hex: string): void => {
		const input = inputs[control];
		if (!input) {
			throw new Error(`no ${control} colour input`);
		}
		input.value = hex;
		input.dispatchEvent(new Event(type, { bubbles: type !== 'blur' }));
		render();
	};
	/** Set the value and fire `input` then `change` in one turn, with no repaint between. */
	const pick = (control: Control, hex: string): void => {
		const input = inputs[control];
		if (!input) {
			throw new Error(`no ${control} colour input`);
		}
		input.value = hex;
		input.dispatchEvent(new Event('input', { bubbles: true }));
		input.dispatchEvent(new Event('change', { bubbles: true }));
		render();
	};
	return {
		ops,
		current,
		fire,
		pick,
		recent: () => store.get().presentationProperties.mruColors ?? [],
		dirty: () => store.get().dirty,
	};
}

type Control = 'fill' | 'line' | 'text' | 'glow';

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
	const ref = REF[control];

	it('shows a dragged shade on the element at once, with no undo step or recent colour', () => {
		const view = mount();
		view.fire(control, 'input', '#330000');

		expect(color(view.current())).toBe('#330000');
		expect(view.ops.canUndo()).toBeFalsy();
		expect(view.dirty()).toBeFalsy();
		expect(view.recent()).toStrictEqual([]);
	});

	it('makes a drag and its settling change ONE undo step and ONE recent colour', () => {
		const view = mount();
		const before = view.current();
		for (const shade of SHADES) {
			view.fire(control, 'input', shade);
			expect(color(view.current())).toBe(shade);
		}
		view.fire(control, 'change', '#990000');

		expect(color(view.current())).toBe('#990000');
		if (control !== 'glow') {
			// A native pick is a custom colour: the theme ref goes.
			expect(ref(view.current())).toBeUndefined();
		}
		expect(view.dirty()).toBeTruthy();
		expect(view.recent()).toStrictEqual(['#990000']);

		view.ops.undo();
		expect(color(view.current())).toBe(color(before));
		expect(ref(view.current())).toStrictEqual(ref(before));
		expect(view.ops.canUndo()).toBeFalsy();
	});

	it('leaves no undo step, and keeps the theme colour, when a drag settles where it started', () => {
		const view = mount();
		const before = view.current();
		view.fire(control, 'input', '#330000');
		view.fire(control, 'input', color(before)!);
		view.fire(control, 'change', color(before)!);

		expect(view.ops.canUndo()).toBeFalsy();
		// The `change` is still a pick: its colour counts as recent, though nothing was edited.
		expect(view.recent()).toStrictEqual([color(before)!.toUpperCase()]);
		expect(color(view.current())).toBe(color(before));
		expect(ref(view.current())).toStrictEqual(ref(before));
	});

	it('commits a pick whose input and change arrive in one turn as one undo step', () => {
		const view = mount();
		const before = view.current();
		view.pick(control, '#00ff00');

		expect(color(view.current())).toBe('#00ff00');
		expect(view.recent()).toStrictEqual(['#00FF00']);
		view.ops.undo();
		expect(color(view.current())).toBe(color(before));
		expect(view.ops.canUndo()).toBeFalsy();
	});

	it('counts a pick of the colour already there as a recent colour, with no undo step', () => {
		const view = mount();
		const before = view.current();
		view.pick(control, color(before)!);

		expect(view.recent()).toStrictEqual([color(before)!.toUpperCase()]);
		expect(view.ops.canUndo()).toBeFalsy();
		expect(ref(view.current())).toStrictEqual(ref(before));
	});

	it('settles a drag that no change event ended when the picker loses focus', () => {
		const view = mount();
		const before = view.current();
		view.fire(control, 'input', '#330000');
		view.fire(control, 'blur', '#330000');

		expect(view.recent()).toStrictEqual(['#330000']);
		view.ops.undo();
		expect(color(view.current())).toBe(color(before));
		expect(view.ops.canUndo()).toBeFalsy();
	});
});
