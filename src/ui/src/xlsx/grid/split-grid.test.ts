// @vitest-environment jsdom
import { createWorkbook, putCell } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { mountGrid } from './index';
import { createTestContext } from './test-context';

let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

function mount(split: { x: number; y: number } | undefined) {
	const wb = createWorkbook();
	const ws = wb.sheets[0]!;
	for (let r = 0; r < 60; r++) putCell(ws, r, 0, { value: `r${r}` });
	if (split) ws.view.split = split;
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	dispose = mountGrid(ctx, container);
	return { ctx, ws, root: ctx.root };
}

const box = (root: ParentNode, name: string) =>
	root.querySelector<HTMLElement>(`.xg-q-${name}`) as HTMLElement;

describe('split window in the grid', () => {
	it('paints four panes and a draggable bar on each axis', () => {
		const { root } = mount({ x: 1500, y: 600 });
		for (const name of ['corner', 'top', 'left', 'main'])
			expect(box(root, name).hidden).toBe(false);
		// 1500 and 600 twentieths of a point are 100 and 40 pixels.
		expect(box(root, 'corner').style.width).toBe('100px');
		expect(box(root, 'corner').style.height).toBe('40px');
		const bars = root.querySelectorAll('.xg-split-bar');
		expect(bars.length).toBe(2);
		expect([...bars].every((bar) => !(bar as HTMLElement).hidden)).toBe(true);
	});

	it('shows the same rows in the top and bottom panes until one is scrolled', () => {
		const { root } = mount({ x: 0, y: 600 });
		expect(box(root, 'corner').hidden).toBe(true);
		const text = (name: string) => box(root, name).textContent ?? '';
		expect(text('top')).toContain('r0');
		expect(text('main')).toContain('r0');
		expect(text('top')).toContain('r1');
	});

	it('removes a bar on double click and the split when the last bar goes', () => {
		const { ws, root } = mount({ x: 1500, y: 600 });
		const dbl = (selector: string) =>
			root.querySelector(selector)!.dispatchEvent(
				new MouseEvent('dblclick', {
					bubbles: true,
					cancelable: true,
					clientX: 500,
					clientY: 500,
				}),
			);
		dbl('.xg-freeze-v');
		expect(ws.view.split).toMatchObject({ x: 0, y: 600 });
		dbl('.xg-freeze-h');
		expect(ws.view.split).toBeUndefined();
	});

	it('keeps a frozen layout free of split bars', () => {
		const { ws, root } = mount(undefined);
		ws.view.freeze = { rows: 1, cols: 1 };
		expect(root.querySelectorAll('.xg-split-bar').length).toBe(0);
	});
});
