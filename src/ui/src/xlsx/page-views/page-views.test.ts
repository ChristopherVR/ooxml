// @vitest-environment jsdom
import {
	createGridMetrics,
	createWorkbook,
	pageBreakPreview,
	pageViewCellRect,
	pageViewLayout,
	putCell,
	type Workbook,
} from 'ooxml-core/xlsx';
import { columnHeaderHeight } from 'ooxml-core/xlsx/ui';
import { afterEach, describe, expect, it } from 'vitest';
import { mountGrid } from '../grid/index';
import { createTestContext } from '../grid/test-context';
import { setViewMode } from '../commands/view';
import { mountPageViews } from './index';

let dispose: (() => void)[] = [];
afterEach(() => {
	for (const fn of dispose) fn();
	dispose = [];
	document.body.replaceChildren();
});
const frame = () => new Promise((resolve) => setTimeout(resolve, 60));

function filled(): Workbook {
	const wb = createWorkbook();
	for (let r = 0; r < 12; r++) putCell(wb.sheets[0]!, r, 0, { value: r + 1 });
	return wb;
}

function setup(wb: Workbook) {
	const ctx = createTestContext(wb);
	const host = document.createElement('div');
	ctx.root.append(host);
	dispose.push(mountGrid(ctx, host));
	const views = mountPageViews(ctx, host, () => 'Book1.xlsx');
	dispose.push(() => views.destroy());
	return { ctx, host };
}

const pointer = (target: Element, type: string, x: number, y: number, init: object = {}) =>
	target.dispatchEvent(
		new MouseEvent(type, {
			bubbles: true,
			cancelable: true,
			clientX: x,
			clientY: y,
			button: 0,
			...init,
		}),
	);

describe('Page Layout view', () => {
	it('shows pages with header and footer placeholders and hides the grid', async () => {
		const wb = filled();
		const { ctx, host } = setup(wb);
		expect(host.dataset.view).toBe('normal');
		setViewMode(ctx, 'pageLayout');
		await frame();
		expect(host.dataset.view).toBe('pageLayout');
		expect(wb.sheets[0]!.view.mode).toBe('pageLayout');
		const pages = ctx.root.querySelectorAll('.xpl-page');
		expect(pages).toHaveLength(1);
		expect(ctx.root.querySelector('.xpl-header .xpl-placeholder')!.textContent).toBe('Add header');
		expect(ctx.root.querySelector('.xpl-footer .xpl-placeholder')!.textContent).toBe('Add footer');
		expect(ctx.root.querySelector('.xpl-inner')!.textContent).toContain('12');
		expect(ctx.root.querySelector('.xpl-colhdr')!.textContent).toBe('A');
		setViewMode(ctx, undefined);
		await frame();
		expect(host.dataset.view).toBe('normal');
		expect(ctx.root.querySelector<HTMLElement>('.xpl-root')!.hidden).toBe(true);
	});

	it('edits a header, saves it to the page setup and renders the field codes', async () => {
		const wb = filled();
		const { ctx } = setup(wb);
		setViewMode(ctx, 'pageLayout');
		await frame();
		const page = pageViewLayout(wb, 0).pages[0]!;
		const root = ctx.root.querySelector('.xpl-root')!;
		pointer(
			root,
			'pointerdown',
			page.header.x + page.header.width / 2 + 24,
			page.header.y + 24 + 3,
		);
		const inputs = ctx.root.querySelectorAll<HTMLInputElement>('.xpl-band-input');
		expect(inputs).toHaveLength(3);
		inputs[1]!.value = 'Page &P of &N';
		inputs[0]!.value = '&A';
		inputs[1]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(wb.sheets[0]!.pageSetup?.header).toBe('&L&A&CPage &P of &N');
		await frame();
		const text = [...ctx.root.querySelectorAll('.xpl-header .xpl-section')].map(
			(n) => n.textContent,
		);
		expect(text).toEqual(['Sheet1', 'Page 1 of 1', '']);
		ctx.session()!.undo();
		await frame();
		expect(ctx.root.querySelector('.xpl-header .xpl-placeholder')).not.toBeNull();
	});

	it('selects and edits a cell, and shows Click to add data on an empty sheet', async () => {
		const wb = filled();
		const { ctx } = setup(wb);
		setViewMode(ctx, 'pageLayout');
		await frame();
		const layout = pageViewLayout(wb, 0);
		const metrics = createGridMetrics(wb.sheets[0]!, { zoom: 100 });
		const rect = pageViewCellRect(layout, metrics, 1, 2, 0)!;
		const root = ctx.root.querySelector('.xpl-root')!;
		pointer(root, 'pointerdown', rect.x + 24 + 3, rect.y + 24 + 3);
		expect(ctx.selection.get().active).toEqual({ row: 2, col: 0 });
		pointer(root, 'dblclick', rect.x + 24 + 3, rect.y + 24 + 3);
		const input = ctx.root.querySelector<HTMLInputElement>('.xpl-edit')!;
		expect(input.value).toBe('3');
		input.value = '42';
		input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(wb.sheets[0]!.rows.get(2)?.get(0)?.value).toBe(42);
		expect(ctx.selection.get().active).toEqual({ row: 3, col: 0 });

		const empty = setup(createWorkbook());
		setViewMode(empty.ctx, 'pageLayout');
		await frame();
		expect(empty.ctx.root.querySelector('.xpl-blank')!.textContent).toBe('Click to add data');
	});
});

describe('Page Break Preview', () => {
	it('paints the break lines and watermarks and moves a break by dragging', async () => {
		const wb = filled();
		wb.sheets[0]!.pageSetup = {
			printArea: { start: { row: 0, col: 0 }, end: { row: 11, col: 2 } },
		};
		const { ctx } = setup(wb);
		ctx.session()!.setPageBreak(0, 'row', 5);
		setViewMode(ctx, 'pageBreakPreview');
		await frame();
		const layer = ctx.root.querySelector<HTMLElement>('.xgpb-layer')!;
		expect(layer.hidden).toBe(false);
		expect(layer.querySelectorAll('.xgpb-mark')).toHaveLength(2);
		expect(layer.querySelector('.xgpb-mark')!.textContent).toBe('Page 1');
		expect(layer.querySelectorAll('.xgpb-line.xgpb-manual')).toHaveLength(1);
		const handle = layer.querySelector<HTMLElement>('.xgpb-grab.xgpb-row:not(:first-of-type)')!;
		expect(handle).not.toBeNull();
		const metrics = createGridMetrics(wb.sheets[0]!, { zoom: 100 });
		const y = columnHeaderHeight(100) + metrics.rowTop(8) + 2;
		const handles = [...layer.querySelectorAll<HTMLElement>('.xgpb-grab.xgpb-row')];
		const line = handles.find(
			(node) =>
				Math.abs(parseFloat(node.style.top) - (columnHeaderHeight(100) + metrics.rowTop(5) - 4)) <
				1,
		)!;
		pointer(line, 'pointerdown', 100, 100);
		pointer(line, 'pointermove', 100, y);
		pointer(line, 'pointerup', 100, y);
		expect(pageBreakPreview(wb, 0).lines.filter((l) => l.axis === 'row')).toEqual([
			{ axis: 'row', before: 8, manual: true },
		]);
		ctx.session()!.undo();
		expect(pageBreakPreview(wb, 0).lines.filter((l) => l.axis === 'row')).toEqual([
			{ axis: 'row', before: 5, manual: true },
		]);
	});

	it('hides the layer in Normal view', async () => {
		const { ctx } = setup(filled());
		await frame();
		expect(ctx.root.querySelector<HTMLElement>('.xgpb-layer')!.hidden).toBe(true);
	});
});
