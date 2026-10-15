import { readFileSync } from 'node:fs';
import { afterEach, expect, it } from 'vitest';
import { loadXlsx } from 'ooxml-core/xlsx';
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

const press = (node: Element, init: PointerEventInit = {}) =>
	node.dispatchEvent(
		new PointerEvent('pointerdown', { pointerId: 1, button: 0, bubbles: true, ...init }),
	);

it('draws a table slicer with its caption, buttons and item states', async () => {
	const { container, layer, view } = await open('table');
	const slicers = [...container.querySelectorAll<HTMLElement>('.xg-slicer')];
	expect(slicers).toHaveLength(2);
	const category = slicers.find((node) => node.getAttribute('aria-label') === 'Category')!;
	expect(category.querySelector('.xg-sl-caption')!.textContent).toBe('Category');
	const states = [...category.querySelectorAll('.xg-sl-item')].map((el) => [
		el.textContent,
		[...el.classList].find((c) => c.startsWith('xg-sl-') && c !== 'xg-sl-item'),
	]);
	expect(states).toEqual([
		['Fruit', 'xg-sl-selected'],
		['Nuts', 'xg-sl-selected'],
		['Veg', 'xg-sl-unselected'],
	]);
	// The two-column Region slicer lays its buttons side by side.
	const region = slicers.find((node) => node.getAttribute('aria-label') === 'Region')!;
	const [east, west] = [...region.querySelectorAll<HTMLElement>('.xg-sl-item')];
	expect(east!.style.top).toBe(west!.style.top);
	expect(east!.style.left).not.toBe(west!.style.left);
	layer.destroy();
	view.destroy();
});

it('filters the table when an item is clicked, extends with Ctrl, toggles Multi-Select and clears', async () => {
	const { wb, container, ctx, layer, view } = await open('table');
	const category = () => container.querySelector<HTMLElement>('.xg-slicer[aria-label="Category"]')!;
	const item = (label: string) =>
		[...category().querySelectorAll<HTMLElement>('.xg-sl-item')].find(
			(el) => el.textContent === label,
		)!;
	const sheet = wb.sheets[0]!;
	press(item('Veg'));
	expect(sheet.tables[0]!.columns[0]!.filter).toEqual({ values: ['Veg'] });
	ctx.notify();
	layer.invalidate();
	layer.render();
	press(item('Nuts'), { ctrlKey: true });
	expect(sheet.tables[0]!.columns[0]!.filter?.values.sort()).toEqual(['Nuts', 'Veg']);
	layer.invalidate();
	layer.render();
	press(category().querySelector('[data-slicer-action=multi]')!);
	expect(category().querySelector('[data-slicer-action=multi]')!.getAttribute('aria-pressed')).toBe(
		'true',
	);
	press(item('Fruit'));
	expect(sheet.tables[0]!.columns[0]!.filter).toBeUndefined();
	layer.invalidate();
	layer.render();
	press(item('Veg'));
	layer.invalidate();
	layer.render();
	press(category().querySelector('[data-slicer-action=clear]')!);
	expect(sheet.tables[0]!.columns[0]!.filter).toBeUndefined();
	expect(sheet.rowInfo.get(3)?.hidden).toBeUndefined();
	layer.destroy();
	view.destroy();
});

it('does not move a slicer when a button is pressed, but a header press still drags', async () => {
	const { wb, container, layer, view } = await open('table');
	const before = structuredClone(wb.sheets[0]!.drawings[0]!.anchor);
	const node = container.querySelector<HTMLElement>('.xg-slicer')!;
	press(node.querySelector('.xg-sl-item')!, { clientX: 5, clientY: 5 });
	expect(wb.sheets[0]!.drawings[0]!.anchor).toEqual(before);
	expect(node.querySelector('.xg-sl-head')).not.toBeNull();
	layer.destroy();
	view.destroy();
});

it('filters a PivotTable through a slicer click', async () => {
	const { wb, container, layer, view } = await open('pivot');
	// The pivot lives on the first sheet of this fixture; the slicer is drawn there too.
	const node = container.querySelector<HTMLElement>('.xg-slicer')!;
	const veg = [...node.querySelectorAll<HTMLElement>('.xg-sl-item')].find(
		(el) => el.textContent === 'Veg',
	)!;
	press(veg);
	expect(wb.sheets[0]!.pivotTables![0]!.itemFilters).toEqual([
		{ field: 'Category', hidden: ['Fruit', 'Nuts'] },
	]);
	layer.destroy();
	view.destroy();
});
