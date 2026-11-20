import { afterEach, expect, it } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { createTestContext } from './test-context';
import { GridView } from './grid-view';
import { DrawingLayer } from './drawings';

afterEach(() => document.body.replaceChildren());
it.each(['pointercancel', 'lostpointercapture', 'Escape'])(
	'restores a drawing without editing its anchor on %s',
	(reason) => {
		const wb = createWorkbook();
		const anchor = {
			from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
			ext: { cx: 952500, cy: 476250 },
		};
		wb.sheets[0]!.drawings.push({ kind: 'unsupported', anchor, description: 'test shape' });
		const ctx = createTestContext(wb);
		const container = document.createElement('div');
		ctx.root.append(container);
		const view = new GridView(ctx, container);
		const layer = new DrawingLayer(view);
		layer.render();
		const node = container.querySelector<HTMLElement>('.xg-obj')!;
		const left = node.style.left;
		const pointer = (type: string, x: number) =>
			node.dispatchEvent(
				new PointerEvent(type, { pointerId: 1, button: 0, clientX: x, clientY: 20, bubbles: true }),
			);
		pointer('pointerdown', 20);
		pointer('pointermove', 60);
		expect(node.style.left).not.toBe(left);
		if (reason === 'Escape')
			document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		else pointer(reason, 60);
		pointer('pointerup', 60);
		expect(node.style.left).toBe(left);
		expect(ctx.workbook()!.sheets[0]!.drawings[0]!.anchor).toEqual(anchor);
		layer.destroy();
		view.destroy();
	},
);

it('uses source crop geometry when painting an Excel picture', () => {
	const wb = createWorkbook();
	wb.sheets[0]!.drawings.push({
		kind: 'image',
		partName: 'xl/media/image1.png',
		contentType: 'image/png',
		anchor: {
			from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
			ext: { cx: 952500, cy: 476250 },
		},
		crop: { cropLeft: 0.25, cropRight: 0.25 },
	});
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	const view = new GridView(ctx, container);
	const layer = new DrawingLayer(view);
	layer.render();
	const img = container.querySelector<HTMLImageElement>('.xg-obj img')!;
	expect(img.style.width).toBe('200px');
	expect(img.style.left).toBe('-50px');
	expect(img.parentElement?.style.overflow).toBe('hidden');
	layer.destroy();
	view.destroy();
});

it('shows an SVG picture from an opaque-origin data URL, never a blob of this origin', () => {
	const wb = createWorkbook();
	const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
	(wb as { source?: unknown }).source = { parts: new Map([['xl/media/image1.svg', svg]]) };
	wb.sheets[0]!.drawings.push({
		kind: 'image',
		partName: 'xl/media/image1.svg',
		contentType: 'image/svg+xml',
		anchor: {
			from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
			ext: { cx: 952500, cy: 476250 },
		},
	});
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	const view = new GridView(ctx, container);
	const layer = new DrawingLayer(view);
	layer.render();
	const img = container.querySelector<HTMLImageElement>('.xg-obj img')!;
	expect(img.getAttribute('src')?.startsWith('data:image/svg+xml;base64,')).toBe(true);
	layer.destroy();
	view.destroy();
});

it('names a chart by its alt text and offers its data to assistive technology', () => {
	const wb = createWorkbook();
	wb.sheets[0]!.drawings.push({
		kind: 'chart',
		chartType: 'column',
		title: 'Sales',
		description: 'Sales doubled in the second week',
		showLegend: false,
		series: [{ name: 'Week', categories: ['One', 'Two'], values: [10, 20] }],
		anchor: {
			from: { row: 0, col: 2, rowOffset: 0, colOffset: 0 },
			ext: { cx: 2857500, cy: 1428750 },
		},
	});
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	const view = new GridView(ctx, container);
	const layer = new DrawingLayer(view);
	layer.render();
	const node = container.querySelector<HTMLElement>('.xg-chart')!;
	expect(node.getAttribute('role')).toBe('img');
	expect(node.getAttribute('aria-label')).toBe('Sales doubled in the second week');
	const describedBy = node.getAttribute('aria-describedby')!;
	expect(node.querySelector(`#${describedBy}`)?.textContent).toBe('Week: One 10; Two 20.');
	layer.destroy();
	view.destroy();
});
