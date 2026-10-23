import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import type { Component } from 'svelte';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import EffectsPanel from './EffectsPanel.svelte';
import FillStrokeSection from './FillStrokeSection.svelte';

/**
 * A live colour drag in the Svelte inspector stays on the element it started
 * on and never overwrites another edit (shared `createLiveColorInput`): the
 * selection moving mid-drag puts the first shape back without landing the pick
 * on the second, an edit to another field of the dragged shape survives the
 * commit and its undo step, and an effects colour field tells a new target
 * from a re-render of its own drag.
 */

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const A_STYLE: ShapeStyle = {
	fillColor: '#4472c4',
	fillColorRef: { scheme: 'accent1' },
	strokeColor: '#ed7d31',
	strokeWidth: 2,
	glowColor: '#ff0000',
	glowRadius: 8,
	glowOpacity: 0.5,
};
const B_STYLE: ShapeStyle = {
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

type Section = Component<{ editor: EditorState; el: PptxElement }>;

function mountDeck(section: Section) {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([
		{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [shape('a', A_STYLE), shape('b', B_STYLE)] },
	]);
	editor.select('a');
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(section, {
		target,
		props: {
			editor,
			get el(): PptxElement {
				return editor.selectedElement!;
			},
		},
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	const style = (id: string): ShapeStyle =>
		(editor.elementById(id) as { shapeStyle: ShapeStyle }).shapeStyle;
	const fire = (selector: string, type: 'input' | 'change' | 'blur', hex: string): void => {
		const input = target.querySelector<HTMLInputElement>(selector);
		if (!input) {
			throw new Error(`no input ${selector}`);
		}
		input.value = hex;
		input.dispatchEvent(new Event(type, { bubbles: type !== 'blur' }));
		flushSync();
	};
	const select = (id: string): void => {
		editor.select(id);
		flushSync();
	};
	return { editor, style, fire, select };
}

const FILL = 'input[type="color"]';

describe('a fill drag the selection leaves', () => {
	it('puts the first shape back and does not land the pick on the second', () => {
		const { editor, style, fire, select } = mountDeck(FillStrokeSection as unknown as Section);
		fire(FILL, 'input', '#330000');
		select('b');
		fire(FILL, 'change', '#330000');

		expect(style('a')).toStrictEqual(A_STYLE);
		expect(style('b')).toStrictEqual(B_STYLE);
		expect(editor.canUndo).toBeFalsy();
		expect(editor.mruColors).toStrictEqual([]);
	});
});

describe('another edit to the dragged shape', () => {
	it('survives the commit, and the one undo step keeps it', () => {
		const { editor, style, fire } = mountDeck(FillStrokeSection as unknown as Section);
		fire(FILL, 'input', '#330000');
		// A collaborator widens the line while the picker is open.
		editor.slides = editor.slides.map((slide) => ({
			...slide,
			elements: slide.elements.map((el) =>
				el.id === 'a'
					? ({ ...el, shapeStyle: { ...style('a'), strokeWidth: 5 } } as PptxElement)
					: el,
			),
		}));
		flushSync();
		fire(FILL, 'change', '#330000');

		expect(style('a').strokeWidth).toBe(5);
		expect(style('a').fillColor).toBe('#330000');
		editor.undo();
		flushSync();
		expect(style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('an effects colour field whose target changes mid-drag', () => {
	it('does not write the first target onto the second', () => {
		const { style, fire, select } = mountDeck(EffectsPanel as unknown as Section);
		const glow = '[data-testid="fx-glow-color"]';
		fire(glow, 'input', '#00ff00');
		fire(glow, 'input', '#ff0000');
		select('b');
		fire(glow, 'blur', '#ff0000');

		expect(style('b')).toStrictEqual(B_STYLE);
		expect(style('a')).toMatchObject({ glowColor: '#ff0000', glowRadius: 8, glowOpacity: 0.5 });
	});
});
