import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import type { Component } from 'svelte';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import EffectsPanel from './EffectsPanel.svelte';
import FillStrokeSection from './FillStrokeSection.svelte';
import TextSection from './TextSection.svelte';

/**
 * The inspector's fill, line and text colour pickers, and the effects colour
 * fields of the generic renderer, preview while the user drags through the
 * native picker, and a drag is ONE undo step and ONE recent colour (shared
 * `createLiveColorInput`), as in React, Vue and vanilla. They used to bind only
 * `onchange`, so nothing moved on the slide until the picker closed.
 *
 * The sections are mounted with an `el` prop that follows the editor's
 * selection, as `InspectorPanel` passes it, so each assertion reads the
 * element the canvas would draw.
 */

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const ACCENT1 = { scheme: 'accent1' } as const;

function shapeEl(): PptxElement {
	return {
		type: 'shape',
		id: 's1',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		text: 'Hi',
		shapeStyle: {
			fillColor: '#4472c4',
			fillColorRef: ACCENT1,
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

function makeEditor(): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [shapeEl()] }]);
	editor.select('s1');
	return editor;
}

function currentEl(editor: EditorState): PptxElement {
	const el = editor.slides[0]?.elements[0];
	if (!el) {
		throw new Error('element missing');
	}
	return el;
}

type Section = Component<{ editor: EditorState; el: PptxElement }>;

/** Mount `section` with `el` following the editor (`live`) or frozen at its first value. */
function mountSection(section: Section, editor: EditorState, live = true): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const first = currentEl(editor);
	const props = {
		editor,
		get el(): PptxElement {
			return live ? currentEl(editor) : first;
		},
	};
	const instance = mount(section, { target, props });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

const style = (el: PptxElement): ShapeStyle => (el as { shapeStyle?: ShapeStyle }).shapeStyle ?? {};
const text = (el: PptxElement): { color?: string; colorRef?: unknown } =>
	(el as { textStyle?: { color?: string; colorRef?: unknown } }).textStyle ?? {};

type Control = 'fill' | 'line' | 'text' | 'glow';

const CONTROLS: Record<
	Control,
	{
		section: Section;
		input(target: HTMLElement): HTMLInputElement | null;
		color(el: PptxElement): string | undefined;
		ref(el: PptxElement): unknown;
	}
> = {
	fill: {
		section: FillStrokeSection as unknown as Section,
		input: (target) => target.querySelectorAll<HTMLInputElement>('input[type="color"]')[0] ?? null,
		color: (el) => style(el).fillColor,
		ref: (el) => style(el).fillColorRef,
	},
	line: {
		section: FillStrokeSection as unknown as Section,
		input: (target) => target.querySelectorAll<HTMLInputElement>('input[type="color"]')[1] ?? null,
		color: (el) => style(el).strokeColor,
		ref: (el) => style(el).strokeColorRef,
	},
	text: {
		section: TextSection as unknown as Section,
		input: (target) => target.querySelector<HTMLInputElement>('input.pptx-svelte-text-color'),
		color: (el) => text(el).color,
		ref: (el) => text(el).colorRef,
	},
	glow: {
		section: EffectsPanel as unknown as Section,
		input: (target) => target.querySelector<HTMLInputElement>('[data-testid="fx-glow-color"]'),
		color: (el) => style(el).glowColor,
		ref: (el) => style(el).glowRadius,
	},
};

function fire(
	input: HTMLInputElement | null,
	type: 'input' | 'change' | 'blur',
	hex: string,
): void {
	if (!input) {
		throw new Error('colour input not found');
	}
	input.value = hex;
	input.dispatchEvent(new Event(type, { bubbles: type !== 'blur' }));
	flushSync();
}

const SHADES = ['#330000', '#660000', '#990000'];

describe.each(['fill', 'line', 'text', 'glow'] as const)('the %s colour picker', (control) => {
	const { section, input, color, ref } = CONTROLS[control];

	it('shows a dragged shade on the element at once, with no undo step or recent colour', () => {
		const editor = makeEditor();
		const target = mountSection(section, editor);
		fire(input(target), 'input', '#330000');

		expect(color(currentEl(editor))).toBe('#330000');
		expect(editor.canUndo).toBeFalsy();
		expect(editor.dirty).toBeFalsy();
		expect(editor.mruColors).toStrictEqual([]);
	});

	it('makes a drag and its settling change ONE undo step and ONE recent colour', () => {
		const editor = makeEditor();
		const before = currentEl(editor);
		const target = mountSection(section, editor);
		for (const shade of SHADES) {
			fire(input(target), 'input', shade);
		}
		expect(color(currentEl(editor))).toBe('#990000');
		fire(input(target), 'change', '#990000');

		expect(color(currentEl(editor))).toBe('#990000');
		expect(editor.dirty).toBeTruthy();
		expect(editor.mruColors).toStrictEqual(['#990000']);

		editor.undo();
		flushSync();
		expect(color(currentEl(editor))).toBe(color(before));
		expect(ref(currentEl(editor))).toStrictEqual(ref(before));
		expect(editor.canUndo).toBeFalsy();
	});

	it('leaves no undo step, and keeps the theme colour, when a drag settles where it started', () => {
		const editor = makeEditor();
		const before = currentEl(editor);
		const target = mountSection(section, editor);
		fire(input(target), 'input', '#330000');
		fire(input(target), 'input', color(before)!);
		fire(input(target), 'change', color(before)!);

		expect(editor.canUndo).toBeFalsy();
		expect(editor.mruColors).toStrictEqual([]);
		expect(color(currentEl(editor))).toBe(color(before));
		expect(ref(currentEl(editor))).toStrictEqual(ref(before));
	});

	it('settles a drag that no change event ended when the picker loses focus', () => {
		const editor = makeEditor();
		const before = currentEl(editor);
		const target = mountSection(section, editor);
		fire(input(target), 'input', '#330000');
		fire(input(target), 'blur', '#330000');

		expect(editor.mruColors).toStrictEqual(['#330000']);
		editor.undo();
		flushSync();
		expect(color(currentEl(editor))).toBe(color(before));
		expect(editor.canUndo).toBeFalsy();
	});

	it('keeps a drag to one undo step while the section still holds the old element', () => {
		const editor = makeEditor();
		const before = currentEl(editor);
		const target = mountSection(section, editor, false);
		for (const shade of SHADES) {
			fire(input(target), 'input', shade);
		}
		expect(color(currentEl(editor))).toBe('#990000');
		fire(input(target), 'change', '#990000');

		editor.undo();
		flushSync();
		expect(color(currentEl(editor))).toBe(color(before));
		expect(editor.canUndo).toBeFalsy();
	});
});
