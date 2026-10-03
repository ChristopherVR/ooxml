// @vitest-environment jsdom
// Regression tests for the core's sheet and selection bookkeeping: per-sheet selection memory,
// one selection formatter, and sheet-change after sheets are inserted, moved or deleted.
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { allCommands } from './commands/index.js';
import { flush, shellFixture } from './test-support/shell';

afterEach(() => document.body.replaceChildren());

const fixture = (sheets = ['One', 'Two', 'Three']) => {
	const fx = shellFixture(createWorkbook({ sheets }));
	fx.core.commands.registerAll(allCommands());
	const of = (type: string) => fx.events.filter((event) => event.type === type);
	return { ...fx, of };
};
const range = (r1: number, c1: number, r2: number, c2: number) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

describe('per-sheet selection memory', () => {
	it('restores each sheet’s last selection when switching tabs', () => {
		const { core } = fixture();
		core.selection.set({ active: { row: 2, col: 1 }, ranges: [range(2, 1, 4, 3)] });
		core.setActiveSheet(1);
		expect(core.selection.get()).toMatchObject({ sheet: 1, active: { row: 0, col: 0 } });
		core.selection.set({ active: { row: 9, col: 9 } });
		core.setActiveSheet(0);
		expect(core.selection.get()).toMatchObject({
			sheet: 0,
			active: { row: 2, col: 1 },
			ranges: [range(2, 1, 4, 3)],
		});
		core.setActiveSheet(1);
		expect(core.selection.get()).toMatchObject({ sheet: 1, active: { row: 9, col: 9 } });
	});
});

describe('one selection formatter', () => {
	it('reports whole columns and rows as A:A and 1:3, once per selection', () => {
		const { core, of } = fixture();
		const before = of('selection-change').length;
		core.selection.set({ active: { row: 0, col: 0 }, ranges: [range(0, 0, 1048575, 0)] });
		expect(of('selection-change').slice(before)).toEqual([
			{ type: 'selection-change', detail: { sheet: 0, ref: 'A:A', active: 'A1' } },
		]);
		core.selection.set({ active: { row: 0, col: 0 }, ranges: [range(0, 0, 2, 16383)] });
		expect(of('selection-change').at(-1)?.detail).toMatchObject({ ref: '1:3' });
		expect(of('selection-change')).toHaveLength(before + 2);
	});
});

describe('sheet structure changes', () => {
	it('announces sheet-change and resets the selection when the active sheet is deleted', async () => {
		const { core, of } = fixture();
		core.setActiveSheet(1);
		core.selection.set({ active: { row: 5, col: 5 } });
		core.setActiveSheet(0);
		core.selection.set({ active: { row: 3, col: 3 } });
		const changes = of('sheet-change').length;
		await core.commands.run('sheet.delete');
		expect(core.workbook!.sheets.map((sheet) => sheet.name)).toEqual(['Two', 'Three']);
		expect(core.activeSheet).toBe(0);
		expect(of('sheet-change').slice(changes)).toEqual([
			{ type: 'sheet-change', detail: { index: 0, name: 'Two' } },
		]);
		expect(core.selection.get()).toMatchObject({ sheet: 0, active: { row: 5, col: 5 } });
		await flush();
	});

	it('switches to the new sheet when one is inserted at the active index', async () => {
		const { core, of } = fixture(['One']);
		core.selection.set({ active: { row: 4, col: 2 } });
		await core.commands.run('sheet.insert');
		expect(core.workbook!.sheets).toHaveLength(2);
		expect(core.activeSheet).toBe(0);
		expect(core.workbook!.sheets[0]!.name).not.toBe('One');
		expect(of('sheet-change').at(-1)?.detail).toMatchObject({ index: 0 });
		expect(core.selection.get()).toMatchObject({ sheet: 0, active: { row: 0, col: 0 } });
		core.setActiveSheet(1);
		expect(core.selection.get()).toMatchObject({ sheet: 1, active: { row: 4, col: 2 } });
	});

	it('follows the shown sheet when a sheet is inserted before it', () => {
		const { core, of } = fixture(['One', 'Two']);
		core.setActiveSheet(1);
		core.session!.addSheet('Zero', 0);
		expect(core.activeSheet).toBe(2);
		expect(core.workbook!.sheets[core.activeSheet]!.name).toBe('Two');
		expect(of('sheet-change').at(-1)?.detail).toEqual({ index: 2, name: 'Two' });
	});
});
