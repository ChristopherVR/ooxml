import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { MAX_COL, parseRange } from '../address';
import { getCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { columnAt } from './columns';
import { createEditSession } from './session';
import type { PasteOptions } from './types';

interface NativeCase extends PasteOptions {
	source: string;
	dest: string;
	error: string | null;
	columns: { width: number; hidden: boolean; storedWidth: number }[];
	value: number;
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-widths.json', import.meta.url), 'utf8'),
) as {
	defaultWidth: number;
	cases: NativeCase[];
};
const range = (ref: string) => parseRange(ref)!;
const setup = () => {
	const workbook = createWorkbook();
	const sheet = workbook.sheets[0]!;
	sheet.defaultColWidth = fixture.defaultWidth;
	sheet.columns = [
		{ min: 0, max: 0, width: 20 },
		{ min: 1, max: 1, width: 30, hidden: true },
	];
	const session = createEditSession(workbook);
	session.setCellValue(0, 0, 0, 2);
	session.setCellValue(0, 0, 1, 3);
	for (let c = 5; c <= 8; c++) {
		sheet.columns.push({ min: c, max: c, width: 12, hidden: c % 2 === 1, outlineLevel: 2 });
		for (let r = 3; r <= 5; r++) session.setCellValue(0, r, c, 9);
	}
	return { workbook, sheet, session };
};

describe('Column Widths paste recorded in Microsoft Excel', () => {
	it('captures metrics beyond the used cells of whole-row copies', () => {
		const { sheet, session } = setup();
		sheet.columns.push({ min: MAX_COL - 1, max: MAX_COL, width: 24 });
		const payload = session.copy(0, range('1:1'));
		expect(payload.cells.cols).toBe(9);
		expect(payload.cells.columnWidths).toHaveLength(MAX_COL + 1);
		expect(payload.cells.columnWidths?.slice(-2)).toEqual([24, 24]);
		expect(payload.cells.columnWidths?.[2]).toBe(fixture.defaultWidth);
	});
	it.each(fixture.cases)('$source -> $dest transpose=$transpose skip=$skipBlanks', (native) => {
		const { sheet, session } = setup();
		const before = structuredClone(sheet.columns);
		const cells = structuredClone(sheet.rows);
		const payload = session.copy(0, range(native.source));
		if (native.error) {
			expect(() =>
				session.paste(0, range(native.dest), payload, {
					mode: 'widths',
					transpose: native.transpose!,
				}),
			).toThrow(RangeError);
			expect(sheet.columns).toEqual(before);
			return;
		}
		session.paste(0, range(native.dest), payload, {
			mode: 'widths',
			transpose: native.transpose!,
			skipBlanks: native.skipBlanks!,
			operation: native.operation!,
		});
		for (let c = 5; c <= 8; c++) {
			const info = columnAt(sheet, c)!;
			const expected = native.columns[c - 5]!;
			expect(!!info.hidden).toBe(expected.hidden);
			expect(info.width).toBe(expected.storedWidth);
			expect(info.hidden ? 0 : info.width).toBe(expected.width);
			expect(info.outlineLevel).toBe(2);
		}
		expect(getCell(sheet, 3, 5)?.value).toBe(native.value);
		expect(sheet.rows).toEqual(cells);
		const after = structuredClone(sheet.columns);
		session.undo();
		expect(sheet.columns).toEqual(before);
		session.redo();
		expect(sheet.columns).toEqual(after);
	});
	it('copies metrics at copy time and uses the receiving workbook default for hidden columns', async () => {
		const { session, sheet } = setup();
		const payload = session.copy(0, range('A1:B1'));
		sheet.columns = [];
		const target = createWorkbook();
		target.sheets[0]!.defaultColWidth = 11;
		createEditSession(target).paste(0, { row: 0, col: 3 }, payload, 'widths');
		const saved = await loadXlsx(await saveXlsx(target));
		expect(columnAt(saved.sheets[0]!, 3)?.width).toBe(20);
		expect(columnAt(saved.sheets[0]!, 4)).toMatchObject({ width: 11, hidden: true });
	});
	it('rejects unavailable metrics, cut payloads and out-of-bounds targets atomically', () => {
		const { sheet, session } = setup();
		const before = structuredClone(sheet.columns);
		const label = session.undoLabel();
		const copy = session.copy(0, range('A1:B1'));
		const cut = session.cut(0, range('A1:B1'));
		const bad = structuredClone(copy);
		bad.cells.columnWidths = [NaN, 10];
		for (const payload of ['1\t2', cut, bad])
			expect(() => session.paste(0, range('F4'), payload, 'widths')).toThrow(RangeError);
		expect(() => session.paste(0, { row: 0, col: MAX_COL }, copy, 'widths')).toThrow(RangeError);
		expect(cut.cut).toBe(true);
		expect(sheet.columns).toEqual(before);
		expect(session.undoLabel()).toBe(label);
	});
});
