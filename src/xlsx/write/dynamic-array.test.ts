import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { createEditSession } from '../edit/session.js';
import { createCalcEngine } from '../formula/engine.js';
import { isSpilledCell } from '../formula/spill.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from './index.js';

const part = async (bytes: Uint8Array, name: string): Promise<string> =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';

function spillingBook() {
	const wb = createWorkbook();
	const s = createEditSession(wb);
	s.setCellInput(0, 0, 0, '=SEQUENCE(3)');
	s.setCellInput(0, 0, 2, '=B1:B2*2');
	s.setCellValue(0, 0, 1, 5);
	s.setCellValue(0, 1, 1, 6);
	s.setCellInput(0, 5, 0, '=A1+1');
	return { wb, s };
}

describe('dynamic arrays on save', () => {
	it('writes spilling formulas as dynamic-array anchors with XLDAPR metadata', async () => {
		const { wb } = spillingBook();
		expect(isSpilledCell(getCell(wb.sheets[0]!, 2, 0))).toBe(true);
		const bytes = await saveXlsx(wb);
		const sheet = await part(bytes, 'xl/worksheets/sheet1.xml');
		expect(sheet).toContain(
			'<c r="A1" cm="1"><f t="array" ref="A1:A3" aca="false">_xlfn.SEQUENCE(3)</f><v>1</v></c>',
		);
		expect(sheet).toContain('<c r="A2"><v>2</v></c>');
		expect(sheet).toContain('<f t="array" ref="C1:C2" aca="false">B1:B2*2</f>');
		expect(sheet).toContain('<c r="A6"><f>A1+1</f><v>2</v></c>');
		const metadata = await part(bytes, 'xl/metadata.xml');
		expect(metadata).toContain('fDynamic="1"');
		expect(metadata).toContain('<rc t="1" v="0"/>');
		expect(await part(bytes, '[Content_Types].xml')).toContain(
			'<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>',
		);
		expect(await part(bytes, 'xl/_rels/workbook.xml.rels')).toContain(
			'relationships/sheetMetadata',
		);
	});

	it('writes no metadata part without dynamic arrays', async () => {
		const wb = createWorkbook();
		createEditSession(wb).setCellInput(0, 0, 0, '=1+1');
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		expect(zip.file('xl/metadata.xml')).toBeNull();
	});

	it('writes a blocked spill as a one-cell dynamic array with #SPILL!', async () => {
		const wb = createWorkbook();
		const s = createEditSession(wb);
		s.setCellValue(0, 1, 0, 'x');
		s.setCellInput(0, 0, 0, '=SEQUENCE(3)');
		expect(getCell(wb.sheets[0]!, 0, 0)?.value).toEqual({ error: '#SPILL!' });
		const sheet = await part(await saveXlsx(wb), 'xl/worksheets/sheet1.xml');
		expect(sheet).toContain('<c r="A1" t="e" cm="1"><f t="array" ref="A1" aca="false">');
	});

	it('drops stale spilled values whose anchor no longer exists', async () => {
		const wb = createWorkbook();
		wb.sheets[0]!.rows.set(
			3,
			new Map([[3, { value: 9, spillAnchor: { row: 0, col: 3 } } as never]]),
		);
		const sheet = await part(await saveXlsx(wb), 'xl/worksheets/sheet1.xml');
		expect(sheet).toContain('<c r="D4"/>');
	});
});

describe('dynamic arrays on load', () => {
	it('reads anchors as dynamic formulas and their cached results as spilled cells', async () => {
		const { wb } = spillingBook();
		const back = await loadXlsx(await saveXlsx(wb));
		const sheet = back.sheets[0]!;
		const anchor = getCell(sheet, 0, 0);
		expect(anchor).toMatchObject({ formula: 'SEQUENCE(3)', dynamicArray: true, value: 1 });
		expect(anchor?.arrayRange).toBeUndefined();
		expect(anchor?.legacyFormula).toBeUndefined();
		expect(isSpilledCell(getCell(sheet, 1, 0))).toBe(true);
		expect(getCell(sheet, 5, 0)?.legacyFormula).toBe(true);
		const calc = createCalcEngine(back);
		calc.recalculateAll();
		expect([0, 1, 2].map((r) => getCell(sheet, r, 0)?.value)).toEqual([1, 2, 3]);
		expect(calc.spillRange(0, 0, 0)).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 0 },
		});
		// Saving again keeps the dynamic array.
		const again = await part(await saveXlsx(back), 'xl/worksheets/sheet1.xml');
		expect(again).toContain('cm="1"><f t="array" ref="A1:A3" aca="false">');
	});

	it('keeps legacy CSE array formulas as fixed array ranges', async () => {
		const wb = createWorkbook();
		wb.sheets[0]!.rows.set(
			0,
			new Map([
				[
					0,
					{
						value: 1,
						formula: 'ROW(A1:A2)',
						arrayRange: { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } },
					},
				],
			]),
		);
		const sheet = await part(await saveXlsx(wb), 'xl/worksheets/sheet1.xml');
		expect(sheet).toContain('<f t="array" ref="A1:A2">ROW(A1:A2)</f>');
		expect(sheet).not.toContain('cm=');
		const back = await loadXlsx(await saveXlsx(wb));
		const cell = getCell(back.sheets[0]!, 0, 0);
		expect(cell?.arrayRange).toBeDefined();
		expect(cell?.dynamicArray).toBeUndefined();
	});
});
