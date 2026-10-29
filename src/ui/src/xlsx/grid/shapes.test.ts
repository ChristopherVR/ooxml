import { readFileSync } from 'node:fs';
import path from 'node:path';
import { afterEach, expect, it } from 'vitest';
import { getValue, loadXlsx } from 'ooxml-core/xlsx';
import { createTestContext } from './test-context';
import { GridView } from './grid-view';
import { DrawingLayer } from './drawings';

afterEach(() => document.body.replaceChildren());

const fixture = () =>
	new Uint8Array(
		readFileSync(
			path.join(
				import.meta.dirname,
				'..',
				'..',
				'..',
				'..',
				'core',
				'xlsx',
				'__fixtures__',
				'excel-shapes.xlsx',
			),
		),
	);

async function mount() {
	const wb = await loadXlsx(fixture());
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	const view = new GridView(ctx, container);
	const layer = new DrawingLayer(view);
	layer.render();
	return { wb, ctx, container, view, layer };
}

const press = (target: Element) =>
	target.dispatchEvent(
		new PointerEvent('pointerdown', { pointerId: 1, button: 0, bubbles: true, cancelable: true }),
	);

it('paints shapes, connectors and groups instead of placeholders', async () => {
	const { container, layer, view } = await mount();
	const shapes = container.querySelectorAll('.xg-shape');
	expect(shapes).toHaveLength(6);
	expect(container.querySelector('.xg-shape svg path')).not.toBeNull();
	expect(container.querySelector('.xg-shape')?.textContent).toContain('Hello');
	expect(container.querySelector('.xg-obj-unsupported')).toBeNull();
	layer.destroy();
	view.destroy();
});

it('paints each form control with its linked state', async () => {
	const { container, layer, view } = await mount();
	const controls = [...container.querySelectorAll<HTMLElement>('.xg-formcontrol')];
	expect(controls.map((c) => c.dataset.controlType)).toEqual([
		'checkbox',
		'radio',
		'radio',
		'dropdown',
		'list',
		'button',
		'spin',
		'scroll',
		'groupBox',
	]);
	expect(controls[0]!.textContent).toContain('Enable');
	expect(controls[3]!.textContent).toContain('Green');
	layer.destroy();
	view.destroy();
});

it('updates the linked cell when a check box, option button or list is used', async () => {
	const { wb, container, layer, view } = await mount();
	const sheet = wb.sheets[0]!;
	const nodes = () => [...container.querySelectorAll<HTMLElement>('.xg-formcontrol')];
	press(nodes()[0]!.querySelector('[data-control-part=toggle]')!);
	expect(getValue(sheet, 0, 2)).toBe(false);
	press(nodes()[0]!.querySelector('[data-control-part=toggle]')!);
	expect(getValue(sheet, 0, 2)).toBe(true);
	press(nodes()[2]!.querySelector('[data-control-part=toggle]')!);
	expect(getValue(sheet, 1, 2)).toBe(2);
	press(nodes()[1]!.querySelector('[data-control-part=toggle]')!);
	expect(getValue(sheet, 1, 2)).toBe(1);
	press(nodes()[4]!.querySelector('[data-control-part=item][data-control-index="3"]')!);
	expect(getValue(sheet, 3, 2)).toBe(3);
	press(nodes()[6]!.querySelector('[data-control-part=step-up]')!);
	expect(getValue(sheet, 4, 2)).toBe(4);
	layer.destroy();
	view.destroy();
});

it('opens the drop-down list and writes the chosen item', async () => {
	const { wb, container, layer, view } = await mount();
	const nodes = () => [...container.querySelectorAll<HTMLElement>('.xg-formcontrol')];
	press(nodes()[3]!.querySelector('[data-control-part=drop]')!);
	const rows = nodes()[3]!.querySelectorAll('.xg-fc-item');
	expect([...rows].map((row) => row.textContent)).toEqual(['Red', 'Green', 'Blue']);
	press(rows[2]!);
	expect(getValue(wb.sheets[0]!, 2, 2)).toBe(3);
	expect(nodes()[3]!.querySelector('.xg-fc-pop')).toBeNull();
	layer.destroy();
	view.destroy();
});
