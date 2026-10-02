// @vitest-environment jsdom
import { createWorkbook, getCell, loadXlsx, saveXlsx } from '@christophervr/xlsx-core';
import { describe, expect, it, vi } from 'vitest';
import { openProtectSheet } from '../dialogs/protect-sheet.js';
import { activeChart, activeTable } from './contextual.js';
import { allCommands } from './index.js';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
	tick,
} from './test-support.js';

function setup(sheets = ['Sheet1']) {
	const ctx = createTestContext(createWorkbook({ sheets }));
	ctx.commands.registerAll(allCommands());
	return ctx;
}

describe('page layout commands', () => {
	it('sets margins, orientation, paper size and the print area, each undoable', async () => {
		const ctx = setup();
		await ctx.commands.run('page.margins', 'narrow');
		await ctx.commands.run('page.orientation-landscape');
		await ctx.commands.run('page.size', 9);
		ctx.select('A1:C5');
		await ctx.commands.run('page.print-area-set');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.pageSetup).toMatchObject({
			orientation: 'landscape',
			paperSize: 9,
			margins: { left: 0.25 },
		});
		expect(ws.pageSetup?.printArea).toEqual({ start: { row: 0, col: 0 }, end: { row: 4, col: 2 } });
		expect(ctx.commands.get('page.orientation-landscape')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('page.print-area-clear');
		expect(ws.pageSetup?.printArea).toBeUndefined();
		ctx.session()!.undo();
		expect(ws.pageSetup?.printArea).toBeDefined();
	});

	it('toggles gridlines and headings for view and print', async () => {
		const ctx = setup();
		const ws = ctx.workbook()!.sheets[0]!;
		await ctx.commands.run('page.gridlines-view');
		expect(ws.view.showGridLines).toBe(false);
		await ctx.commands.run('page.headings-view');
		expect(ws.view.showHeaders).toBe(false);
		await ctx.commands.run('page.gridlines-print');
		await ctx.commands.run('page.headings-print');
		expect(ws.printOptions).toEqual({ gridLines: true, headings: true });
		expect(ctx.session()!.undoLabel()).toBe('Print options');
		await ctx.commands.run('page.gridlines-print');
		expect(ws.printOptions).toEqual({ headings: true });
	});
});

describe('review and view commands', () => {
	it('walks comments, deletes them and protects the sheet', async () => {
		const ctx = setup();
		const session = ctx.session()!;
		session.setComment(0, { row: 4, col: 1 }, 'b', 'me');
		session.setComment(0, { row: 1, col: 2 }, 'a', 'me');
		await ctx.commands.run('review.next-comment');
		expect(ctx.selection.get().active).toEqual({ row: 1, col: 2 });
		await ctx.commands.run('review.next-comment');
		expect(ctx.selection.get().active).toEqual({ row: 4, col: 1 });
		await ctx.commands.run('review.previous-comment');
		expect(ctx.selection.get().active).toEqual({ row: 1, col: 2 });
		await ctx.commands.run('review.delete-comment');
		expect(ctx.workbook()!.sheets[0]!.comments).toHaveLength(1);
		ctx.workbook()!.sheets[0]!.protection = { sheet: true };
		expect(ctx.commands.get('review.protect-sheet')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('review.protect-sheet');
		expect(ctx.workbook()!.sheets[0]!.protection).toBeUndefined();
	});

	it('protects a sheet with a password and asks for it to unprotect', async () => {
		const ctx = setup();
		ctx.dialogs.register('protect-sheet', (c) => openProtectSheet(c));
		void ctx.commands.run('review.protect-sheet');
		await tick();
		const dialog = dialogEl(ctx, 'protect-sheet');
		setValue(inputByLabel(dialog, 'Password to unprotect sheet:'), 'secret');
		setValue(inputByLabel(dialog, 'Reenter password to proceed:'), 'secret');
		clickButton(dialog, 'OK');
		await tick();
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.protection?.passwordHash).toBeTruthy();
		expect(ctx.session()!.undoLabel()).toBe('Protect sheet');
		void ctx.commands.run('review.protect-sheet');
		await tick();
		const ask = dialogEl(ctx, 'unprotect-sheet');
		setValue(inputByLabel(ask, 'Password:'), 'wrong');
		clickButton(ask, 'OK');
		await tick();
		expect(ws.protection?.sheet).toBe(true);
		setValue(inputByLabel(ask, 'Password:'), 'secret');
		clickButton(ask, 'OK');
		await tick();
		expect(ctx.workbook()!.sheets[0]!.protection).toBeUndefined();
	});

	it('locks the workbook structure as an undoable step', async () => {
		const ctx = setup();
		void ctx.commands.run('review.protect-workbook');
		await tick();
		clickButton(dialogEl(ctx, 'protect-workbook'), 'OK');
		await tick();
		expect(ctx.workbook()!.structureLocked).toBe(true);
		expect(ctx.session()!.undoLabel()).toBe('Protect workbook');
		ctx.session()!.undo();
		expect(ctx.workbook()!.structureLocked).toBeUndefined();
	});

	it('freezes panes, zooms through the grid and calculates', async () => {
		const setZoom = vi.fn();
		const ctx = createTestContext(createWorkbook(), { grid: { setZoom } });
		ctx.commands.registerAll(allCommands());
		ctx.select('C4');
		await ctx.commands.run('view.freeze-panes');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.view.freeze).toEqual({ rows: 3, cols: 2 });
		await ctx.commands.run('view.freeze-panes');
		expect(ws.view.freeze).toBeUndefined();
		await ctx.commands.run('view.freeze-top-row');
		expect(ws.view.freeze).toEqual({ rows: 1, cols: 0 });
		await ctx.commands.run('view.unfreeze');
		expect(ws.view.freeze).toBeUndefined();
		await ctx.commands.run('view.zoom-100');
		expect(setZoom).toHaveBeenCalledWith(100);
		ctx.select('A1:B2');
		await ctx.commands.run('view.zoom-selection');
		expect(setZoom).toHaveBeenLastCalledWith(400);
		await ctx.commands.run('formulas.show-formulas');
		expect(ws.view.showFormulas).toBe(true);
		ctx.session()!.setCellInput(0, 0, 0, '=1+1');
		expect(await ctx.commands.run('formulas.calculate-now')).toBe(true);
		expect(getCell(ws, 0, 0)?.value).toBe(2);
		// Manual calculation on the live session: edits stop recalculating until F9.
		await ctx.commands.run('formulas.calc-manual');
		expect(ctx.workbook()!.calcMode).toBe('manual');
		expect(ctx.commands.get('formulas.calc-manual')?.checked?.(ctx)).toBe(true);
		ctx.session()!.setCellInput(0, 1, 0, '5');
		ctx.session()!.setCellInput(0, 0, 0, '=A2*2');
		await ctx.commands.run('formulas.calculate-sheet');
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)?.value).toBe(10);
		await ctx.commands.run('formulas.calc-automatic');
		expect(ctx.workbook()!.calcMode).toBeUndefined();
	});
});

describe('contextual commands', () => {
	it('shows Table Design inside a table and edits its options and total row', async () => {
		const ctx = setup();
		const session = ctx.session()!;
		[
			['Item', 'Qty'],
			['a', '1'],
			['b', '2'],
		].forEach((row, r) => row.forEach((v, c) => session.setCellInput(0, r, c, v)));
		session.createTable(0, { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }, true);
		ctx.select('A2');
		expect(activeTable(ctx)?.name).toBe('Table1');
		await ctx.commands.run('table.banded-columns');
		expect(activeTable(ctx)?.showColumnStripes).toBe(true);
		await ctx.commands.run('table.style', 'TableStyleLight9');
		expect(activeTable(ctx)?.styleName).toBe('TableStyleLight9');
		await ctx.commands.run('table.total-row');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.tables[0]?.range.end.row).toBe(3);
		expect(getCell(ws, 3, 0)?.value).toBe('Total');
		expect(getCell(ws, 3, 1)?.value).toBe(3);
		session.undo();
		// A workbook-wide undo step restores fresh sheet objects.
		expect(ctx.workbook()!.sheets[0]!.tables[0]?.range.end.row).toBe(2);
		ctx.select('E10');
		expect(activeTable(ctx)).toBeUndefined();
		expect(ctx.commands.isEnabled('table.total-row')).toBe(false);
	});

	it('follows the selected chart and toggles its legend', async () => {
		const ctx = setup();
		const ws = ctx.workbook()!.sheets[0]!;
		ws.drawings.push({
			kind: 'chart',
			chartType: 'column',
			series: [],
			showLegend: true,
			anchor: {
				from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
				to: { row: 10, col: 6, rowOffset: 0, colOffset: 0 },
			},
			partName: 'xl/charts/chart1.xml',
		});
		ctx.select('C3');
		expect(activeChart(ctx)).toBeUndefined();
		ctx.selection.set({ drawing: 0 });
		expect(activeChart(ctx)?.index).toBe(0);
		await ctx.commands.run('chart.legend');
		const chart = ws.drawings[0];
		expect(chart?.kind === 'chart' && chart.showLegend).toBe(false);
		// The loaded chart part is kept (the core patches it on save) so its detail survives.
		expect(chart?.kind === 'chart' && chart.partName).toBe('xl/charts/chart1.xml');
		expect(ctx.session()!.undoLabel()).toBe('Edit chart');
		await ctx.commands.run('chart.delete');
		expect(ws.drawings).toHaveLength(0);
	});
});

describe('core-backed edits', () => {
	it('follows the selected drawing over the active cell', () => {
		const ctx = setup();
		const ws = ctx.workbook()!.sheets[0]!;
		ws.drawings.push({
			kind: 'chart',
			chartType: 'pie',
			series: [],
			showLegend: true,
			anchor: {
				from: { row: 20, col: 20, rowOffset: 0, colOffset: 0 },
				ext: { cx: 9525, cy: 9525 },
			},
		});
		ctx.select('A1');
		expect(activeChart(ctx)).toBeUndefined();
		ctx.selection.set({ drawing: 0 });
		expect(activeChart(ctx)?.chart.chartType).toBe('pie');
		ctx.select('B2');
		expect(activeChart(ctx)).toBeUndefined();
	});

	it('turns a table header row off and on', async () => {
		const ctx = setup();
		const session = ctx.session()!;
		[
			['Item', 'Qty'],
			['a', '1'],
		].forEach((row, r) => row.forEach((v, c) => session.setCellInput(0, r, c, v)));
		session.createTable(0, { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }, true);
		ctx.select('A2');
		await ctx.commands.run('table.header-row');
		const table = () => ctx.workbook()!.sheets[0]!.tables[0];
		expect(table()?.headerRow).toBe(false);
		expect(table()?.range.start.row).toBe(1);
		expect(session.undoLabel()).toBe('Table options');
		await ctx.commands.run('table.header-row');
		expect(table()?.headerRow).toBe(true);
	});

	it('groups whole columns and rows, asking for a cell range', async () => {
		const ctx = setup();
		ctx.dialogs.register('group', async () => 'cols');
		const ws = ctx.workbook()!.sheets[0]!;
		const grouped = (col: number) =>
			ws.columns.some((c) => c.min <= col && c.max >= col && (c.outlineLevel ?? 0) > 0);
		ctx.select('B:C');
		await ctx.commands.run('data.group');
		expect(grouped(1) && grouped(2)).toBe(true);
		ctx.select('2:3');
		await ctx.commands.run('data.group');
		expect(ws.rowInfo.get(1)?.outlineLevel).toBe(1);
		ctx.select('E5:F6');
		await ctx.commands.run('data.group');
		expect(grouped(4)).toBe(true);
		ctx.select('B:C');
		expect(ctx.commands.isEnabled('data.ungroup')).toBe(true);
		await ctx.commands.run('data.ungroup');
		expect(grouped(1)).toBe(false);
	});

	it('applies a built-in cell style by name', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, 'x');
		await ctx.commands.run('home.cell-styles', 'Good');
		const wb = ctx.workbook()!;
		const style = wb.styles[getCell(wb.sheets[0]!, 0, 0)?.styleId ?? 0];
		expect(style?.cellStyleName).toBe('Good');
	});

	it('print options and page setup survive a save', async () => {
		const ctx = setup();
		await ctx.commands.run('page.gridlines-print');
		await ctx.commands.run('page.orientation-landscape');
		expect(ctx.session()!.undoLabel()).toBe('Page setup');
		const loaded = await loadXlsx(await saveXlsx(ctx.workbook()!));
		expect(loaded.sheets[0]!.pageSetup?.orientation).toBe('landscape');
		expect(loaded.sheets[0]!.printOptions?.gridLines).toBe(true);
	});
});
