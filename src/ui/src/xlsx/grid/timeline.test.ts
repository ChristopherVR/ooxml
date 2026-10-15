import { readFileSync } from 'node:fs';
import { afterEach, expect, it } from 'vitest';
import { getCell, loadXlsx } from 'ooxml-core/xlsx';
import { createTestContext } from './test-context';
import { GridView } from './grid-view';
import { DrawingLayer } from './drawings';

afterEach(() => document.body.replaceChildren());

const fixture = (name: string) =>
	new Uint8Array(readFileSync(`../core/xlsx/__fixtures__/excel-slicer-${name}.xlsx`));

async function open(kind: string) {
	const wb = await loadXlsx(fixture(kind));
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	const view = new GridView(ctx, container);
	const layer = new DrawingLayer(view);
	layer.render();
	return { wb, ctx, view, layer, container };
}

const pointer = (node: Element, type: string, init: PointerEventInit = {}) =>
	node.dispatchEvent(new PointerEvent(type, { pointerId: 1, button: 0, bubbles: true, ...init }));
const cell = (container: HTMLElement, label: string) =>
	[...container.querySelectorAll<HTMLElement>('.xg-tl-cell')].find(
		(el) => el.textContent === label,
	)!;

it('draws a timeline with its caption, selection label, level list and selected months', async () => {
	const { container, layer, view } = await open('timeline-range');
	const node = container.querySelector<HTMLElement>('.xg-timeline')!;
	expect(node.querySelector('.xg-tl-caption')!.textContent).toBe('Day');
	expect(node.querySelector('.xg-tl-selection')!.textContent).toBe('Feb 2025 - Apr 2025');
	expect(node.querySelector<HTMLSelectElement>('.xg-tl-level')!.value).toBe('2');
	const on = [...node.querySelectorAll('.xg-tl-on')].map((el) => el.textContent);
	expect(on).toEqual(['Feb', 'Mar', 'Apr']);
	expect(node.querySelector('.xg-tl-group')!.textContent).toBe('2025');
	layer.destroy();
	view.destroy();
});

it('selects a period by clicking, extends with Shift and clears again', async () => {
	const { wb, container, layer, view } = await open('timeline');
	const sheet = wb.sheets[0]!;
	const total = () => {
		for (let row = 2; row < 10; row++)
			if (getCell(sheet, row, 0)?.value === 'Grand Total') return getCell(sheet, row, 1)?.value;
		return undefined;
	};
	expect(total()).toBe(66);
	const click = (label: string, init: PointerEventInit = {}) => {
		const target = cell(container, label);
		pointer(target, 'pointerdown', init);
		pointer(target, 'pointerup', init);
		layer.invalidate();
		layer.render();
	};
	click('Feb');
	expect(wb.slicers!.timelineCaches[0]!.selection).toEqual({
		start: '2025-02-01T00:00:00',
		end: '2025-02-28T00:00:00',
	});
	expect(total()).toBe(20);
	click('Apr', { shiftKey: true });
	expect(wb.slicers!.timelineCaches[0]!.selection!.end).toBe('2025-04-30T00:00:00');
	expect(total()).toBe(40);
	pointer(container.querySelector('[data-timeline-action=clear]')!, 'pointerdown');
	expect(wb.slicers!.timelineCaches[0]!.selection).toBeUndefined();
	expect(total()).toBe(66);
	layer.destroy();
	view.destroy();
});

it('changes the time level from the level list', async () => {
	const { wb, container, layer, view } = await open('timeline');
	const list = container.querySelector<HTMLSelectElement>('.xg-tl-level')!;
	list.value = '1';
	list.dispatchEvent(new Event('change', { bubbles: true }));
	expect(wb.slicers!.timelines[0]!.level).toBe(1);
	layer.invalidate();
	layer.render();
	expect([...container.querySelectorAll('.xg-tl-cell')].map((el) => el.textContent)).toEqual([
		'Q1',
		'Q2',
		'Q3',
		'Q4',
	]);
	layer.destroy();
	view.destroy();
});
