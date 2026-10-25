/**
 * A live colour drag in the vanilla inspector stays on the element it started
 * on and never overwrites another edit (shared `createLiveColorInput`): the
 * selection moving mid-drag puts the first shape back without landing the pick
 * on the second, an edit to another field of the dragged shape survives the
 * commit and its undo step, an effects colour field tells a new target from a
 * re-render of its own drag, and a pick with several shapes selected edits the
 * primary one alone, as `applyToSelected` always has. Wired end to end: the
 * sections, the chrome's inspector handlers, the real edit actions and history
 * (`colour-drag.test-support.ts`).
 */
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { mountInspectorEditor } from './colour-drag.test-support';
import { createFillSection } from './fill-section';
import type { InspectorState } from './types';

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

function mount() {
	const { store, ops, handlers } = mountInspectorEditor(
		[shape('a', A_STYLE), shape('b', B_STYLE)],
		{ selectedElementId: 'a', selectedElementIds: ['a'] },
	);
	const fill = createFillSection(
		document,
		(key) => key,
		() => document.createElement('section'),
		handlers,
	);
	const element = (id: string) =>
		store.get().slides[0]!.elements.find((el) => el.id === id) as PptxElement & {
			shapeStyle: ShapeStyle;
		};
	/** Repaint the section from the store, as the viewer does after each change. */
	const render = (): void => {
		const id = store.get().selectedElementId!;
		const el = element(id);
		fill.update({
			hasSelection: true,
			canShape: true,
			elementId: id,
			fillColor: el.shapeStyle.fillColor,
			strokeColor: el.shapeStyle.strokeColor,
			shapeStyle: el.shapeStyle,
			fillOpacity: 1,
			gradientEnabled: false,
			gradient: { type: 'linear', angle: 90, stops: [] },
			recentColors: [],
		} as unknown as InspectorState);
	};
	render();
	const fillInput = fill.el.querySelector<HTMLInputElement>(
		':scope > [role="group"] input[type="color"]',
	)!;
	const glowInput = fill.el.querySelector<HTMLInputElement>('[data-testid="fx-glow-color"]')!;
	const fire = (input: HTMLInputElement, type: 'input' | 'change' | 'blur', hex: string): void => {
		input.value = hex;
		input.dispatchEvent(new Event(type, { bubbles: type !== 'blur' }));
		render();
	};
	const select = (...ids: string[]): void => {
		store.set({ selectedElementId: ids[0]!, selectedElementIds: ids });
		render();
	};
	return {
		ops,
		store,
		fillInput,
		glowInput,
		fire,
		select,
		style: (id: string) => element(id).shapeStyle,
		recent: () => store.get().presentationProperties.mruColors ?? [],
	};
}

describe('a fill drag the selection leaves', () => {
	it('puts the first shape back and does not land the pick on the second', () => {
		const view = mount();
		view.fire(view.fillInput, 'input', '#330000');
		view.select('b');
		view.fire(view.fillInput, 'change', '#330000');

		expect(view.style('a')).toStrictEqual(A_STYLE);
		expect(view.style('b')).toStrictEqual(B_STYLE);
		expect(view.ops.canUndo()).toBeFalsy();
		expect(view.recent()).toStrictEqual([]);
	});
});

describe('a fill drag the selection leaves after another edit to the shape', () => {
	it('puts the colour back on the shape as it is now, keeping that edit', () => {
		const view = mount();
		view.fire(view.fillInput, 'input', '#330000');
		const state = view.store.get();
		view.store.set({
			slides: state.slides.map((slide) => ({
				...slide,
				elements: slide.elements.map((el) =>
					el.id === 'a'
						? ({ ...el, shapeStyle: { ...view.style('a'), strokeWidth: 5 } } as PptxElement)
						: el,
				),
			})),
		});
		view.select('b');
		view.fire(view.fillInput, 'change', '#330000');

		expect(view.style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('another edit to the dragged shape', () => {
	it('survives the commit, and the one undo step keeps it', () => {
		const view = mount();
		view.fire(view.fillInput, 'input', '#330000');
		// A collaborator widens the line while the picker is open.
		const state = view.store.get();
		view.store.set({
			slides: state.slides.map((slide) => ({
				...slide,
				elements: slide.elements.map((el) =>
					el.id === 'a'
						? ({ ...el, shapeStyle: { ...view.style('a'), strokeWidth: 5 } } as PptxElement)
						: el,
				),
			})),
		});
		view.fire(view.fillInput, 'change', '#330000');

		expect(view.style('a').strokeWidth).toBe(5);
		expect(view.style('a').fillColor).toBe('#330000');
		view.ops.undo();
		expect(view.style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('an effects colour field whose target changes mid-drag', () => {
	it('does not write the first target onto the second', () => {
		const view = mount();
		view.fire(view.glowInput, 'input', '#00ff00');
		view.fire(view.glowInput, 'input', '#ff0000');
		view.select('b');
		view.fire(view.glowInput, 'blur', '#ff0000');

		expect(view.style('b')).toStrictEqual(B_STYLE);
		expect(view.style('a')).toMatchObject({
			glowColor: '#ff0000',
			glowRadius: 8,
			glowOpacity: 0.5,
		});
	});

	it('puts the first target back as it is now, keeping an edit made there since', () => {
		const view = mount();
		view.fire(view.glowInput, 'input', '#00ff00');
		const state = view.store.get();
		view.store.set({
			slides: state.slides.map((slide) => ({
				...slide,
				elements: slide.elements.map((el) =>
					el.id === 'a'
						? ({ ...el, shapeStyle: { ...view.style('a'), glowRadius: 12 } } as PptxElement)
						: el,
				),
			})),
		});
		view.select('b');
		view.fire(view.glowInput, 'blur', '#00ff00');

		expect(view.style('a')).toStrictEqual({ ...A_STYLE, glowRadius: 12 });
		expect(view.style('b')).toStrictEqual(B_STYLE);
	});
});

describe('a pick with several shapes selected', () => {
	it('edits the primary shape alone, and copies nothing onto the others', () => {
		const view = mount();
		view.select('a', 'b');
		view.fire(view.fillInput, 'input', '#00ff00');
		view.fire(view.fillInput, 'change', '#00ff00');

		expect(view.style('a').fillColor).toBe('#00ff00');
		expect(view.style('b')).toStrictEqual(B_STYLE);
	});
});
