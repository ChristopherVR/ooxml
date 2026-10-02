import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { describe, expect, it, vi } from 'vitest';
import { autoSumRange, fillSource } from './editing.js';
import { allCommands } from './index.js';
import { findSpecial } from './select-special.js';
import { createTestContext } from './test-support.js';
import { target } from './util.js';

function setup(sheets = ['Sheet1']) {
	const ctx = createTestContext(createWorkbook({ sheets }));
	ctx.commands.registerAll(allCommands());
	return ctx;
}
const value = (ctx: ReturnType<typeof setup>, row: number, col: number, sheet = 0) =>
	getCell(ctx.workbook()!.sheets[sheet]!, row, col)?.value ?? null;

function fill(ctx: ReturnType<typeof setup>, rows: (string | number)[][]) {
	rows.forEach((cells, r) =>
		cells.forEach((v, c) => ctx.session()!.setCellInput(0, r, c, String(v))),
	);
}

describe('cells commands', () => {
	it('inserts and deletes rows and columns for the selection', async () => {
		const ctx = setup();
		fill(ctx, [[1], [2], [3]]);
		ctx.select('A2:B3');
		await ctx.commands.run('cells.insert-rows');
		expect([value(ctx, 0, 0), value(ctx, 1, 0), value(ctx, 3, 0)]).toEqual([1, null, 2]);
		await ctx.commands.run('cells.delete-rows');
		expect(value(ctx, 1, 0)).toBe(2);
		ctx.select('A1');
		await ctx.commands.run('cells.insert-columns');
		expect(value(ctx, 0, 1)).toBe(1);
		await ctx.commands.run('cells.delete-columns');
		expect(value(ctx, 0, 0)).toBe(1);
	});

	it('inserts, hides and deletes sheets, respecting structure protection', async () => {
		const ctx = setup(['Sheet1']);
		expect(ctx.commands.isEnabled('sheet.delete')).toBe(false);
		await ctx.commands.run('sheet.insert');
		const wb = ctx.workbook()!;
		expect(wb.sheets).toHaveLength(2);
		expect(ctx.activeSheet()).toBe(0);
		expect(ctx.commands.isEnabled('sheet.hide')).toBe(true);
		await ctx.commands.run('sheet.hide');
		expect(wb.sheets[0]!.state).toBe('hidden');
		expect(ctx.activeSheet()).toBe(1);
		expect(ctx.commands.isEnabled('sheet.unhide')).toBe(true);
		wb.structureLocked = true;
		expect(ctx.commands.isEnabled('sheet.insert')).toBe(false);
		delete wb.structureLocked;
		ctx.session()!.setSheetState(0, 'visible');
		await ctx.commands.run('sheet.delete');
		expect(wb.sheets).toHaveLength(1);
	});

	it('hides rows and columns, autofits with the grid measure and toggles cell lock', async () => {
		const measureText = vi.fn((text: string) => text.length * 10);
		const ctx = createTestContext(createWorkbook(), { grid: { measureText } });
		ctx.commands.registerAll(allCommands());
		ctx.session()!.setCellInput(0, 0, 0, 'A rather long label');
		ctx.select('A1:B1');
		await ctx.commands.run('format.autofit-columns');
		expect(measureText).toHaveBeenCalled();
		expect(ctx.workbook()!.sheets[0]!.columns[0]?.width).toBeGreaterThan(20);
		await ctx.commands.run('format.hide-rows');
		expect(ctx.workbook()!.sheets[0]!.rowInfo.get(0)?.hidden).toBe(true);
		await ctx.commands.run('format.unhide-rows');
		expect(ctx.workbook()!.sheets[0]!.rowInfo.get(0)?.hidden).toBeUndefined();
		expect(ctx.commands.get('format.lock-cell')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('format.lock-cell');
		expect(ctx.commands.get('format.lock-cell')?.checked?.(ctx)).toBe(false);
	});
});

describe('editing commands', () => {
	it('guesses the AutoSum range above, else to the left', () => {
		const ctx = setup();
		fill(ctx, [[1, 2, 3], [4], [5]]);
		const t = target(ctx)!;
		expect(autoSumRange(t, 3, 0)).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } });
		expect(autoSumRange(t, 0, 3)).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } });
		expect(autoSumRange(t, 5, 5)).toBeUndefined();
	});

	it('AutoSum starts editing with the guess, or totals a block', async () => {
		const beginEdit = vi.fn();
		const ctx = createTestContext(createWorkbook(), { grid: { beginEdit } });
		ctx.commands.registerAll(allCommands());
		fill(ctx as never, [
			[1, 2],
			[3, 4],
		]);
		ctx.select('A3');
		await ctx.commands.run('home.autosum');
		expect(beginEdit).toHaveBeenCalledWith('=SUM(A1:A2)');
		ctx.select('A1:B2');
		await ctx.commands.run('home.autosum-max');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(getCell(ws, 2, 0)?.formula).toBe('MAX(A1:A2)');
		expect(getCell(ws, 2, 1)?.value).toBe(4);
	});

	it('fills down, right, up and left', async () => {
		const ctx = setup();
		fill(ctx, [[1, 'x']]);
		ctx.select('A1:B3');
		await ctx.commands.run('home.fill-down');
		expect([value(ctx, 2, 0), value(ctx, 2, 1)]).toEqual([1, 'x']);
		expect(
			fillSource({ start: { row: 4, col: 0 }, end: { row: 4, col: 1 } }, 'down')?.source.start.row,
		).toBe(3);
		expect(
			fillSource({ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }, 'down'),
		).toBeUndefined();
		ctx.select('A1:C1');
		await ctx.commands.run('home.fill-right');
		expect(value(ctx, 0, 2)).toBe(1);
	});

	it('clears contents, formats and everything', async () => {
		const ctx = setup();
		fill(ctx, [[1, 2]]);
		ctx.select('A1:B1');
		await ctx.commands.run('home.bold');
		await ctx.commands.run('home.clear-contents');
		expect(value(ctx, 0, 0)).toBeNull();
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)?.styleId).toBeTruthy();
		await ctx.commands.run('home.clear-formats');
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)).toBeUndefined();
	});

	it('sorts the current region with its header and filters', async () => {
		const ctx = setup();
		fill(ctx, [
			['Name', 'Score'],
			['b', 2],
			['c', 3],
			['a', 1],
		]);
		ctx.select('B2');
		await ctx.commands.run('data.sort-desc');
		expect([value(ctx, 0, 0), value(ctx, 1, 0), value(ctx, 3, 0)]).toEqual(['Name', 'c', 'a']);
		await ctx.commands.run('data.filter');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.autoFilter?.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 3, col: 1 } });
		expect(ctx.commands.get('data.filter')?.checked?.(ctx)).toBe(true);
		ctx.session()!.filterColumn(0, 0, ['a']);
		expect(ws.rowInfo.get(1)?.hidden).toBe(true);
		await ctx.commands.run('data.filter-clear');
		expect(ws.rowInfo.get(1)?.hidden).toBeUndefined();
		ctx.select('A1');
		await ctx.commands.run('data.sort-asc');
		expect(value(ctx, 1, 0)).toBe('a');
		await ctx.commands.run('data.filter');
		expect(ws.autoFilter).toBeUndefined();
	});

	it('selects special cells', async () => {
		const ctx = setup();
		fill(ctx, [
			[1, '=A1*2'],
			['', 'text'],
		]);
		ctx.session()!.setComment(0, { row: 1, col: 1 }, 'note', 'me');
		expect(findSpecial(ctx, { kind: 'formulas' })).toEqual([
			{ start: { row: 0, col: 1 }, end: { row: 0, col: 1 } },
		]);
		expect(findSpecial(ctx, { kind: 'constants', types: ['numbers'] })).toEqual([
			{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
		]);
		expect(findSpecial(ctx, { kind: 'blanks' })).toEqual([
			{ start: { row: 1, col: 0 }, end: { row: 1, col: 0 } },
		]);
		expect(findSpecial(ctx, { kind: 'lastCell' })).toEqual([
			{ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
		]);
		await ctx.commands.run('home.select-comments');
		expect(ctx.selection.get().active).toEqual({ row: 1, col: 1 });
		await ctx.commands.run('home.select-validation');
		expect(ctx.toasts.at(-1)?.message).toBe('No cells were found.');
	});
});
