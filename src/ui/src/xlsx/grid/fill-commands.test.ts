import { afterEach, describe, expect, it } from 'vitest';
import { getCell, parseRange } from 'ooxml-core/xlsx';
import { createTestContext } from './test-context';
import { gridCommands, type CommandHost } from './grid-commands';

const range = (ref: string) => {
	const result = parseRange(ref);
	if (!result) throw new Error(ref);
	return result;
};
// These tests invoke only the fill commands; clipboard and menu methods are unused.
const host = {} as CommandHost;

afterEach(() => document.body.replaceChildren());

describe('Fill Down and Fill Right commands', () => {
	it.each([
		['edit.fill-down', 'A1:A3', 2, 0],
		['edit.fill-right', 'A1:C1', 0, 2],
	] as const)('%s repeats weekday text and makes one undo step', async (id, ref, row, col) => {
		const ctx = createTestContext();
		const session = ctx.session()!;
		const sheet = session.workbook.sheets[0]!;
		session.setCellValue(0, 0, 0, 'Monday');
		ctx.selection.set({ ranges: [range(ref)] });
		ctx.commands.registerAll(gridCommands(host));
		expect(await ctx.commands.run(id)).toBe(true);
		expect(getCell(sheet, row, col)?.value).toBe('Monday');
		session.undo();
		expect(getCell(sheet, row, col)).toBeUndefined();
		expect(getCell(sheet, 0, 0)?.value).toBe('Monday');
		session.redo();
		expect(getCell(sheet, row, col)?.value).toBe('Monday');
	});

	it('fills a single row from the row above and honors read-only mode', async () => {
		const ctx = createTestContext();
		const session = ctx.session()!;
		const sheet = session.workbook.sheets[0]!;
		session.setCellValue(0, 0, 0, 'Monday');
		ctx.selection.set({ ranges: [range('A2')] });
		ctx.commands.registerAll(gridCommands(host));
		ctx.setReadOnly(true);
		expect(await ctx.commands.run('edit.fill-down')).toBe(false);
		expect(getCell(sheet, 1, 0)).toBeUndefined();
		ctx.setReadOnly(false);
		expect(await ctx.commands.run('edit.fill-down')).toBe(true);
		expect(getCell(sheet, 1, 0)?.value).toBe('Monday');
	});
});
