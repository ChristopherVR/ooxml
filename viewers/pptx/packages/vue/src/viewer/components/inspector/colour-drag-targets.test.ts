/**
 * A live colour drag in the Vue inspector stays on the element it started on
 * and never overwrites another edit (shared `createLiveColorInput`): the
 * selection moving mid-drag puts the first shape back without landing the pick
 * on the second, an edit to another field of the dragged shape survives the
 * commit and its undo step, an effects colour field tells a new target from a
 * re-render of its own drag, and a drag on a group member or a template
 * element is put back once the selection has left it. Wired as
 * `PowerPointViewer` wires it (`colour-drag.test-support.ts`).
 */
import { disableAutoUnmount, enableAutoUnmount } from '@vue/test-utils';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { afterAll, afterEach, describe, expect, it } from 'vitest';
import type { Component } from 'vue';
import { nextTick } from 'vue';

import type { Deck } from './colour-drag.test-support';
import { mountColourPanel } from './colour-drag.test-support';
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

/** `panel` over shapes `a` and `b` (or `deck`), with `a` selected unless `deck` says otherwise. */
function harness(panel: Component, selector: string, deck: Partial<Deck> = {}) {
	const view = mountColourPanel(panel, selector, {
		...deck,
		elements: deck.elements ?? [shape('a', A_STYLE), shape('b', B_STYLE)],
	});
	const style = (id: string): ShapeStyle =>
		(
			findIn(
				[...view.slides.value[0]!.elements, ...(view.templateElementsBySlideId.value.s ?? [])],
				id,
			) as { shapeStyle: ShapeStyle }
		).shapeStyle;
	return { ...view, style };
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
