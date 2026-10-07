import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { cellError, isErrorCode, type CellValue } from '../model';
import { styleAt } from '../styles';
import { createWorkbook } from '../workbook';
import { parseRange } from '../address';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';
import type { PasteOptions } from './types';

interface NativeCase extends PasteOptions {
	source: string;
	destination: string;
	value: CellValue;
	formula: string | null;
	bold?: boolean;
	italic?: boolean;
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-arithmetic.json', import.meta.url), 'utf8'),
) as { cases: NativeCase[] };

describe('paste arithmetic recorded in Microsoft Excel', () => {
	it('tiles transposed operations, skips blanks, translates formulas and round trips', async () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		const range = (ref: string) => parseRange(ref)!;
		const sheet = workbook.sheets[0]!;
		session.setCellInput(0, 1, 0, '=B2+$B$1');
		session.setCellValue(0, 0, 1, 1);
		session.setCellValue(0, 1, 1, 2);
		const payload = session.copy(0, range('A1:A2'));
		session.setRangeValues(0, { row: 3, col: 2 }, [
			[10, 10, 10, 10],
			[10, 10, 10, 10],
		]);
		session.paste(0, range('C4:F5'), payload, {
			operation: 'add',
			transpose: true,
			skipBlanks: true,
		});
		expect(getCell(sheet, 3, 2)?.value).toBe(10);
		expect(getCell(sheet, 3, 3)?.formula).toBe('10+(D5+$B$1)');
		expect(getCell(sheet, 4, 5)?.formula).toBe('10+(F6+$B$1)');
		expect(getCell(sheet, 3, 3)?.value).toBe(22);
		session.undo();
		expect(getCell(sheet, 3, 3)?.value).toBe(10);
		session.redo();
		const loaded = await loadXlsx(await saveXlsx(workbook));
		expect(getCell(loaded.sheets[0]!, 3, 3)?.formula).toBe('10+(D5+$B$1)');
		expect(getCell(loaded.sheets[0]!, 3, 3)?.value).toBe(22);
	});
	it('rejects operations for cut cells without changing the move payload', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		session.setCellValue(0, 0, 0, 2);
		session.setCellValue(0, 0, 2, 10);
		const payload = session.cut(0, parseRange('A1')!);
		const label = session.undoLabel();
		expect(() => session.paste(0, { row: 0, col: 2 }, payload, { operation: 'add' })).toThrow(
			RangeError,
		);
		expect(payload.cut).toBe(true);
		expect(session.undoLabel()).toBe(label);
		expect(getCell(workbook.sheets[0]!, 0, 2)?.value).toBe(10);
	});
	it.each(fixture.cases)('$mode $operation: $destination with $source', (native) => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		const sheet = workbook.sheets[0]!;
		for (const [col, input] of [
			[0, native.source],
			[2, native.destination],
		] as const)
			if (isErrorCode(input)) session.setCellValue(0, 0, col, cellError(input));
			else session.setCellInput(0, 0, col, input);
		session.applyStyle(0, [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }], {
			font: { bold: true },
		});
		session.applyStyle(0, [{ start: { row: 0, col: 2 }, end: { row: 0, col: 2 } }], {
			font: { italic: true },
		});
		const before = structuredClone(getCell(sheet, 0, 2));
		const payload = session.copy(0, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
		session.paste(0, { row: 0, col: 2 }, payload, native);
		const result = getCell(sheet, 0, 2);
		expect(result?.value ?? null).toEqual(native.value);
		expect(result?.formula ?? null).toBe(native.formula);
		if (native.bold !== undefined) {
			const style = styleAt(workbook, result?.styleId);
			expect(style.font.bold ?? false).toBe(native.bold);
			expect(style.font.italic ?? false).toBe(native.italic);
		}
		session.undo();
		expect(getCell(sheet, 0, 2)).toEqual(before);
		session.redo();
		expect(getCell(sheet, 0, 2)?.value ?? null).toEqual(native.value);
	});
});
