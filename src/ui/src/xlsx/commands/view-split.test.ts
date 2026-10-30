// @vitest-environment jsdom
import { createWorkbook } from 'ooxml-core/xlsx';
import { describe, expect, it } from 'vitest';
import { allCommands } from './index';
import { createTestContext } from './test-support';

// The grid measures the painted window; the command only needs its answer.
const splitAt = (active: { row: number; col: number }) => ({
	x: active.col * 960,
	y: active.row * 300,
	activePane: 'bottomRight' as const,
});

function setup() {
	const ctx = createTestContext(createWorkbook(), { grid: { splitAt } });
	ctx.commands.registerAll(allCommands());
	return ctx;
}

describe('View > Split', () => {
	it('toggles a split at the active cell and removes it again', async () => {
		const ctx = setup();
		const ws = ctx.workbook()!.sheets[0]!;
		ctx.select('C4');
		await ctx.commands.run('view.split');
		expect(ws.view.split).toEqual({ x: 1920, y: 900, activePane: 'bottomRight' });
		expect(ctx.commands.get('view.split')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('view.split');
		expect(ws.view.split).toBeUndefined();
	});

	it('undoes as one step', async () => {
		const ctx = setup();
		const ws = ctx.workbook()!.sheets[0]!;
		ctx.select('B2');
		await ctx.commands.run('view.split');
		expect(ctx.session()!.undoLabel()).toBe('Split window');
		ctx.session()!.undo();
		expect(ws.view.split).toBeUndefined();
	});

	it('excludes Freeze Panes in both directions', async () => {
		const ctx = setup();
		const ws = ctx.workbook()!.sheets[0]!;
		ctx.select('C4');
		await ctx.commands.run('view.freeze-panes');
		expect(ws.view.freeze).toEqual({ rows: 3, cols: 2 });
		await ctx.commands.run('view.split');
		expect(ws.view.freeze).toBeUndefined();
		expect(ws.view.split).toBeDefined();
		await ctx.commands.run('view.freeze-top-row');
		expect(ws.view.split).toBeUndefined();
		expect(ws.view.freeze).toEqual({ rows: 1, cols: 0 });
	});
});
