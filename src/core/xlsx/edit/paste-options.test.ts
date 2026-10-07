import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { parseRange } from '../address';
import { styleAt } from '../styles';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { resolvePasteOptions } from './paste-options';
import { createEditSession } from './session';
import type { PasteOptions } from './types';

interface NativeCell {
	value: unknown;
	formula: string | null;
	bold: boolean;
	italic: boolean;
	fill: string | null;
}
interface NativeCase extends Required<Omit<PasteOptions, 'operation'>> {
	cells: NativeCell[];
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-options.json', import.meta.url), 'utf8'),
) as { version: string; build: number; cases: NativeCase[] };
const range = (ref: string) => {
	const result = parseRange(ref);
	if (!result) throw new Error(ref);
	return result;
};
const setup = () => {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	const sheet = workbook.sheets[0]!;
	const at = { row: 0, col: 2 };
	session.applyStyle(0, [range('A1')], {
		font: { bold: true },
		fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FF0000' } },
	});
	session.setCellInput(0, 1, 0, '=""');
	session.setCellValue(0, 2, 0, 0);
	session.setCellValue(0, 3, 0, false);
	const payload = session.copy(0, range('A1:A5'));
	return { workbook, session, sheet, at, payload };
};

describe('paste options recorded in Microsoft Excel', () => {
	it('records every mode, transpose and skip-blanks combination', () => {
		expect(fixture.cases).toHaveLength(16);
		expect(new Set(fixture.cases.map((c) => `${c.mode}/${c.transpose}/${c.skipBlanks}`)).size).toBe(
			16,
		);
	});
	it.each(fixture.cases)('$mode transpose=$transpose skipBlanks=$skipBlanks', (native) => {
		const { workbook, session, sheet, at, payload } = setup();
		const dest = range(native.transpose ? 'C1:G1' : 'C1:C5');
		for (let i = 0; i < 5; i++)
			session.setCellValue(0, native.transpose ? 0 : i, native.transpose ? i + 2 : 2, 9);
		session.applyStyle(0, [dest], {
			font: { bold: false, italic: true },
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: '00FF00' } },
		});
		const read = (): NativeCell[] =>
			Array.from({ length: 5 }, (_, i) => {
				const cell = getCell(sheet, native.transpose ? 0 : i, native.transpose ? i + 2 : 2);
				const style = styleAt(workbook, cell?.styleId);
				return {
					value: cell?.value ?? null,
					formula: cell?.formula ?? null,
					bold: style.font.bold ?? false,
					italic: style.font.italic ?? false,
					fill:
						style.fill.type === 'pattern' && style.fill.pattern === 'solid'
							? (style.fill.fgColor?.rgb?.slice(-6) ?? null)
							: null,
				};
			});
		const before = read();
		session.paste(0, at, payload, native);
		expect(read()).toEqual(native.cells);
		session.undo();
		expect(read()).toEqual(before);
		session.redo();
		expect(read()).toEqual(native.cells);
	});

	it('retains skip-blanks results through a save and reload', async () => {
		const { workbook, session, at, payload } = setup();
		session.setRangeValues(0, at, [[9], [9], [9], [9], [9]]);
		session.paste(0, at, payload, { mode: 'values', skipBlanks: true });
		const loaded = await loadXlsx(await saveXlsx(workbook));
		expect([0, 1, 2, 3, 4].map((row) => getCell(loaded.sheets[0]!, row, 2)?.value ?? null)).toEqual(
			[9, '', 0, false, 9],
		);
	});
	it('rejects invalid option values', () => {
		for (const invalid of [
			null,
			'bogus',
			{ mode: 'bogus' },
			{ skipBlanks: 'false' },
			{ transpose: 1 },
			{ operation: 'bogus' },
		])
			expect(() => resolvePasteOptions(invalid)).toThrow(RangeError);
		expect(resolvePasteOptions('transpose')).toEqual({
			mode: 'all',
			transpose: true,
			skipBlanks: false,
			operation: 'none',
		});
	});
	it('tiles transposed values while preserving skipped destinations', () => {
		const { session, sheet, payload } = setup();
		session.setRangeValues(0, { row: 6, col: 2 }, [Array(10).fill(9), Array(10).fill(9)]);
		session.paste(0, range('C7:L8'), payload, {
			mode: 'values',
			transpose: true,
			skipBlanks: true,
		});
		for (const row of [6, 7])
			expect(Array.from({ length: 10 }, (_, i) => getCell(sheet, row, i + 2)?.value)).toEqual([
				9,
				'',
				0,
				false,
				9,
				9,
				'',
				0,
				false,
				9,
			]);
		session.undo();
		expect(getCell(sheet, 7, 3)?.value).toBe(9);
	});
	it('rejects skip blanks for a cut without consuming the move or changing cells', () => {
		const { session, sheet, at } = setup();
		const payload = session.cut(0, range('A1:A5'));
		session.setCellValue(0, at.row, at.col, 9);
		const label = session.undoLabel();
		expect(() => session.paste(0, at, payload, { skipBlanks: true })).toThrow(RangeError);
		expect(payload.cut).toBe(true);
		expect(getCell(sheet, 2, 0)?.value).toBe(0);
		expect(getCell(sheet, at.row, at.col)?.value).toBe(9);
		expect(session.undoLabel()).toBe(label);
	});
});
