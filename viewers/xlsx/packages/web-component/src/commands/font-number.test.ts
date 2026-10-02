import { createWorkbook, getCell, styleAt, type Workbook } from '@christophervr/xlsx-core';
import { describe, expect, it } from 'vitest';
import { allCommands } from './index.js';
import { createTestContext } from './test-support.js';

function setup(workbook: Workbook = createWorkbook()) {
	const ctx = createTestContext(workbook);
	ctx.commands.registerAll(allCommands());
	return ctx;
}
const styleOf = (ctx: ReturnType<typeof setup>, row: number, col: number) => {
	const wb = ctx.workbook()!;
	return styleAt(wb, getCell(wb.sheets[0]!, row, col)?.styleId);
};

describe('font commands', () => {
	it('toggles bold on the selection and reports the checked state', async () => {
		const ctx = setup();
		ctx.select('A1:B2');
		expect(ctx.commands.get('home.bold')?.checked?.(ctx)).toBe(false);
		expect(await ctx.commands.run('home.bold')).toBe(true);
		expect(styleOf(ctx, 1, 1).font.bold).toBe(true);
		expect(ctx.commands.get('home.bold')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('home.bold');
		expect(styleOf(ctx, 0, 0).font.bold).toBeUndefined();
	});

	it('sets font name and size, and grows / shrinks through the size list', async () => {
		const ctx = setup();
		await ctx.commands.run('home.font-name', 'Arial');
		await ctx.commands.run('home.font-size', '14');
		expect(styleOf(ctx, 0, 0).font).toMatchObject({ name: 'Arial', size: 14 });
		expect(ctx.commands.get('home.font-name')?.value?.(ctx)).toBe('Arial');
		await ctx.commands.run('home.grow-font');
		expect(styleOf(ctx, 0, 0).font.size).toBe(16);
		await ctx.commands.run('home.shrink-font');
		await ctx.commands.run('home.shrink-font');
		expect(styleOf(ctx, 0, 0).font.size).toBe(12);
		await ctx.commands.run('home.font-size', 'abc');
		expect(styleOf(ctx, 0, 0).font.size).toBe(12);
	});

	it('applies underline variants, strikethrough, colours and clears them', async () => {
		const ctx = setup();
		await ctx.commands.run('home.underline-double');
		expect(styleOf(ctx, 0, 0).font.underline).toBe('double');
		await ctx.commands.run('home.underline');
		expect(styleOf(ctx, 0, 0).font.underline).toBe('single');
		await ctx.commands.run('home.strikethrough');
		expect(styleOf(ctx, 0, 0).font.strike).toBe(true);
		await ctx.commands.run('home.fill-color', { rgb: 'FFFF00' });
		expect(styleOf(ctx, 0, 0).fill).toEqual({
			type: 'pattern',
			pattern: 'solid',
			fgColor: { rgb: 'FFFF00' },
		});
		await ctx.commands.run('home.fill-color', undefined);
		expect(styleOf(ctx, 0, 0).fill).toEqual({ type: 'pattern', pattern: 'none' });
		await ctx.commands.run('home.font-color', { theme: 5, tint: 0.4 });
		expect(styleOf(ctx, 0, 0).font.color).toEqual({ theme: 5, tint: 0.4 });
	});

	it('applies border presets as one undo step and remembers the last one', async () => {
		const ctx = setup();
		ctx.select('A1:B2');
		await ctx.commands.run('home.borders', 'outside');
		expect(styleOf(ctx, 0, 0).border.top?.style).toBe('thin');
		expect(styleOf(ctx, 0, 0).border.bottom).toBeUndefined();
		expect(styleOf(ctx, 1, 1).border.right?.style).toBe('thin');
		await ctx.commands.run('home.borders', { preset: 'all', style: 'thick' });
		expect(styleOf(ctx, 0, 0).border.bottom?.style).toBe('thick');
		await ctx.commands.run('home.borders', 'none');
		await ctx.commands.run('home.borders');
		expect(styleOf(ctx, 0, 0).border).toEqual({});
		// The last call re-applied 'none' (a no-op on the cells), so two undo steps back to thick.
		ctx.session()!.undo();
		ctx.session()!.undo();
		expect(styleOf(ctx, 0, 0).border.bottom?.style).toBe('thick');
	});

	it('is disabled in read-only mode and on a protected sheet', async () => {
		const ctx = setup();
		ctx.setReadOnly(true);
		expect(ctx.commands.isEnabled('home.bold')).toBe(false);
		expect(await ctx.commands.run('home.bold')).toBe(false);
		ctx.setReadOnly(false);
		ctx.workbook()!.sheets[0]!.protection = { sheet: true };
		expect(ctx.commands.isEnabled('home.bold')).toBe(false);
		ctx.workbook()!.sheets[0]!.protection = { sheet: true, allow: ['formatCells'] };
		expect(ctx.commands.isEnabled('home.bold')).toBe(true);
	});
});

describe('alignment commands', () => {
	it('sets and toggles alignment, wrap, orientation and indent', async () => {
		const ctx = setup();
		await ctx.commands.run('home.align-center');
		expect(styleOf(ctx, 0, 0).alignment?.horizontal).toBe('center');
		expect(ctx.commands.get('home.align-center')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('home.align-center');
		expect(styleOf(ctx, 0, 0).alignment?.horizontal).toBeUndefined();
		await ctx.commands.run('home.align-top');
		await ctx.commands.run('home.wrap-text');
		await ctx.commands.run('home.orientation', 90);
		await ctx.commands.run('home.indent-increase');
		expect(styleOf(ctx, 0, 0).alignment).toMatchObject({
			vertical: 'top',
			wrapText: true,
			textRotation: 90,
			indent: 1,
			horizontal: 'left',
		});
		await ctx.commands.run('home.indent-decrease');
		expect(styleOf(ctx, 0, 0).alignment?.indent).toBeUndefined();
	});

	it('merges and centers, then unmerges with the same toggle', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, 'Title');
		ctx.select('A1:C1');
		await ctx.commands.run('home.merge-center');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.merges).toHaveLength(1);
		expect(ctx.commands.get('home.merge-center')?.checked?.(ctx)).toBe(true);
		await ctx.commands.run('home.merge-center');
		expect(ws.merges).toHaveLength(0);
		ctx.select('A1:C2');
		await ctx.commands.run('home.merge-across');
		expect(ws.merges).toHaveLength(2);
		expect(ctx.commands.isEnabled('home.unmerge')).toBe(true);
		await ctx.commands.run('home.unmerge');
		expect(ws.merges).toHaveLength(0);
	});
});

describe('number commands', () => {
	it('applies presets, percent, comma, accounting and steps decimals', async () => {
		const ctx = setup();
		ctx.session()!.setCellValue(0, 0, 0, 1.5);
		await ctx.commands.run('home.number-format', 'percentage');
		expect(styleOf(ctx, 0, 0).numFmt).toBe('0.00%');
		await ctx.commands.run('home.decimal-decrease');
		expect(styleOf(ctx, 0, 0).numFmt).toBe('0.0%');
		await ctx.commands.run('home.decimal-increase');
		await ctx.commands.run('home.decimal-increase');
		expect(styleOf(ctx, 0, 0).numFmt).toBe('0.000%');
		await ctx.commands.run('home.percent');
		expect(styleOf(ctx, 0, 0).numFmt).toBe('0%');
		await ctx.commands.run('home.accounting', 'eur');
		expect(styleOf(ctx, 0, 0).numFmt).toContain('€');
		await ctx.commands.run('home.number-format', 'General');
		await ctx.commands.run('home.decimal-increase');
		expect(styleOf(ctx, 0, 0).numFmt).toBe('0.00');
		await ctx.commands.run('home.format-currency');
		expect(styleOf(ctx, 0, 0).numFmt).toBe('"$"#,##0.00_);\\("$"#,##0.00\\)');
	});
});
