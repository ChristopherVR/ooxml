// @vitest-environment jsdom
// What a screen reader gets from the grid: rows of cells it can explore as a table, the active
// cell as the sink's active descendant, and a label that names the cell's context and content.
import { createEditSession, createWorkbook, putCell, type Workbook } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountGrid } from './index';
import { createTestContext } from './test-context';

let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

const frame = () => new Promise((resolve) => setTimeout(resolve, 40));

function salesBook(): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	putCell(sheet, 0, 0, { value: 'Quarter' });
	putCell(sheet, 0, 1, { value: 'Sales' });
	putCell(sheet, 1, 0, { value: 'Q1' });
	putCell(sheet, 1, 1, { value: 100 });
	putCell(sheet, 2, 0, { value: 'Q2' });
	putCell(sheet, 2, 1, { value: 200, formula: 'B2*2' });
	return wb;
}

async function setup(workbook = salesBook()) {
	const ctx = createTestContext(workbook);
	const container = document.createElement('div');
	ctx.root.append(container);
	dispose = mountGrid(ctx, container);
	await frame();
	const sink = ctx.root.querySelector<HTMLTextAreaElement>('.xg-cell-field textarea')!;
	const grid = ctx.root.querySelector<HTMLElement>('[role="grid"]')!;
	const key = (k: string, init: KeyboardEventInit = {}) =>
		sink.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...init }));
	const activeCell = () => {
		const id = sink.getAttribute('aria-activedescendant');
		return id ? ctx.root.getElementById(id) : null;
	};
	const cell = (text: string) =>
		[...grid.querySelectorAll('[role="gridcell"]')].find((node) => node.textContent === text);
	return { ctx, sink, grid, key, activeCell, cell };
}

describe('the accessible grid', () => {
	it('exposes the in-cell picture alt text through the active grid cell', async () => {
		const workbook = createWorkbook();
		const png = Uint8Array.from(
			atob(
				'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNoTgkFAAKqAT2gS4eLAAAAAElFTkSuQmCC',
			),
			(char) => char.charCodeAt(0),
		);
		createEditSession(workbook).addCellPicture(0, { row: 0, col: 0 }, png, 'Logo');
		const { activeCell, sink } = await setup(workbook);
		expect(activeCell()?.textContent).toBe('Logo');
		expect(sink.getAttribute('aria-label')).toBe('A1, Logo');
	});

	it('exposes the shown cells as grid cells in rows, with sheet positions', async () => {
		const { ctx, grid, cell } = await setup();
		expect(grid.getAttribute('aria-label')).toBe('Spreadsheet grid');
		// Indices are the sheet's row and column numbers, as Excel reports them: B2 is row 2, column 2.
		expect(grid.getAttribute('aria-rowcount')).toBe('1048576');
		expect(grid.getAttribute('aria-colcount')).toBe('16384');
		const hundred = cell('100')!;
		expect(hundred.getAttribute('aria-colindex')).toBe('2');
		expect(hundred.parentElement!.getAttribute('aria-rowindex')).toBe('2');
		expect(
			hundred.parentElement!.querySelector('[role="rowheader"]')?.hasAttribute('aria-colindex'),
		).toBe(false);
		expect(hundred.parentElement!.querySelector('[role="rowheader"]')?.textContent).toBe('2');
		const columnB = [...grid.querySelectorAll('[role="columnheader"]')].find(
			(node) => node.textContent === 'B',
		);
		expect(columnB?.getAttribute('aria-colindex')).toBe('2');
		expect(columnB?.parentElement?.hasAttribute('aria-rowindex')).toBe(false);
		// Every cell and header sits in a row, and nothing painted is exposed outside the grid.
		const owned = ctx.root.querySelectorAll(
			'[role="gridcell"],[role="columnheader"],[role="rowheader"]',
		);
		expect([...owned].every((node) => node.parentElement?.getAttribute('role') === 'row')).toBe(
			true,
		);
		expect([...owned].every((node) => grid.contains(node))).toBe(true);
		for (const painted of ctx.root.querySelectorAll('.xg-cells, .xg-hbox'))
			expect(painted.closest('[aria-hidden="true"]')).not.toBeNull();
	});

	it('points the focus sink at the active cell and follows the selection', async () => {
		const { sink, key, activeCell } = await setup();
		expect(activeCell()?.textContent).toBe('Quarter');
		expect(sink.getAttribute('aria-label')).toBe('A1, Quarter');
		key('ArrowDown');
		key('ArrowRight');
		await frame();
		expect(activeCell()?.textContent).toBe('100');
		expect(activeCell()?.getAttribute('aria-selected')).toBe('true');
		expect(sink.getAttribute('aria-label')).toBe('B2, 100');
	});

	it('hands focus back to the editor while a cell is edited', async () => {
		const { sink, key, activeCell } = await setup();
		sink.value = 'x';
		sink.dispatchEvent(new Event('input', { bubbles: true }));
		expect(sink.hasAttribute('aria-activedescendant')).toBe(false);
		key('Escape');
		expect(activeCell()?.textContent).toBe('Quarter');
	});

	it('announces the table column, the formula and the comment of the active cell', async () => {
		const wb = salesBook();
		const sheet = wb.sheets[0]!;
		sheet.tables.push({
			id: 1,
			name: 'SalesTable',
			displayName: 'SalesTable',
			range: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			headerRow: true,
			totalsRow: false,
			columns: [{ name: 'Quarter' }, { name: 'Sales' }],
		});
		sheet.comments.push({ address: { row: 2, col: 1 }, author: 'Ann', text: 'Check the total' });
		const { ctx, sink } = await setup(wb);
		ctx.selection.set({
			active: { row: 2, col: 1 },
			anchor: { row: 2, col: 1 },
			ranges: [{ start: { row: 2, col: 1 }, end: { row: 2, col: 1 } }],
		});
		expect(sink.getAttribute('aria-label')).toBe(
			'Sales, B3, 200. Formula: =B2*2. Comment: Check the total',
		);
	});

	it('keeps the label right when the cell changes under an unchanged selection', async () => {
		const { ctx, sink, activeCell } = await setup();
		ctx.session()!.setCellValue(0, 0, 0, 'Year');
		// Repaints without a selection change rebuild the layer once they settle.
		await vi.waitFor(
			() => {
				expect(sink.getAttribute('aria-label')).toBe('A1, Year');
				expect(activeCell()?.textContent).toBe('Year');
			},
			{ timeout: 3000, interval: 50 },
		);
	});

	it('catches up on changes made while the editor was removed once it is back', async () => {
		const { ctx, activeCell } = await setup();
		ctx.host.remove();
		ctx.session()!.setCellValue(0, 0, 0, 'Year');
		await frame();
		document.body.append(ctx.host);
		ctx.grid()!.invalidate();
		await vi.waitFor(() => expect(activeCell()?.textContent).toBe('Year'), {
			timeout: 3000,
			interval: 50,
		});
	});

	it('spans a merged cell and gives its covered cells no cell of their own', async () => {
		const wb = salesBook();
		wb.sheets[0]!.merges.push({ start: { row: 4, col: 0 }, end: { row: 4, col: 2 } });
		putCell(wb.sheets[0]!, 4, 0, { value: 'Total' });
		const { grid, cell } = await setup(wb);
		expect(cell('Total')?.getAttribute('aria-colspan')).toBe('3');
		const row5 = grid.querySelector('[aria-rowindex="5"]')!;
		const indices = [...row5.querySelectorAll('[role="gridcell"]')].map((node) =>
			node.getAttribute('aria-colindex'),
		);
		expect(indices.slice(0, 2)).toEqual(['1', '4']);
	});
});
