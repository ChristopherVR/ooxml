// @vitest-environment jsdom
import { createWorkbook, getCell, serialToDate } from '@christophervr/xlsx-core';
import { describe, expect, it } from 'vitest';
import { allCommands } from './index.js';
import { nowSerial } from './insert-now.js';
import { createTestContext } from './test-support.js';

describe('insert date and time', () => {
	it('computes local wall-clock serials', () => {
		const at = new Date(2026, 9, 3, 14, 30, 45);
		expect(serialToDate(nowSerial('date', at)).toISOString()).toBe('2026-10-03T00:00:00.000Z');
		expect(nowSerial('time', at)).toBeCloseTo((14 * 60 + 30) / 1440, 9);
	});

	it('puts the date in the active cell with a date format, as one undo step', async () => {
		const ctx = createTestContext(createWorkbook());
		ctx.commands.registerAll(allCommands());
		ctx.select('B2');
		await ctx.commands.run('edit.insert-date');
		const wb = ctx.workbook()!;
		const cell = getCell(wb.sheets[0]!, 1, 1);
		expect(typeof cell?.value).toBe('number');
		expect(wb.styles[cell?.styleId ?? 0]?.numFmt).toBe('m/d/yyyy');
		expect(ctx.session()!.undoLabel()).toBe('Insert date');
		await ctx.commands.run('edit.insert-time');
		expect(wb.styles[getCell(wb.sheets[0]!, 1, 1)?.styleId ?? 0]?.numFmt).toBe('m/d/yyyy');
	});
});
