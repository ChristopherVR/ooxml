import { parseAddress, parseRange } from '../xlsx/address';
import { createEditSession } from '../xlsx/edit/session';
import type { Workbook } from '../xlsx/model';
import { internStyle, styleAt } from '../xlsx/styles';
import { createWorkbook } from '../xlsx/workbook';
import { saveXlsx } from '../xlsx/write/save';
import { findXlsxCells, inspectXlsx, readXlsxRange, readXlsxTable } from './xlsx';

/** Puts a value with a number format straight into the model, as a loaded file would hold it. */
function put(workbook: Workbook, sheet: number, ref: string, value: number, numFmt: string) {
	const at = parseAddress(ref)!;
	const styleId = internStyle(workbook, { ...styleAt(workbook, 0), numFmt });
	const target = workbook.sheets[sheet]!;
	const row = target.rows.get(at.row) ?? new Map();
	target.rows.set(at.row, row.set(at.col, { value, styleId }));
}

/** A sales sheet with a table and formatted values, a hidden lookup sheet and an empty one. */
async function sample(): Promise<Uint8Array> {
	const workbook = createWorkbook({ sheets: ['Sales', 'Lookup', 'Empty'] });
	const session = createEditSession(workbook);
	const inputs: [string, string][] = [
		['A1', 'Region'],
		['B1', 'Opened'],
		['C1', 'Share'],
		['D1', 'Amount'],
		['A2', 'North'],
		['B2', '4/3/2024'],
		['C2', '15.3%'],
		['D2', '$1,234.50'],
		['A3', 'South'],
		['B3', '2024-04-03 10:30'],
		['C3', '0.5'],
		['D3', '=D2*2'],
		['A4', 'TRUE'],
		['B4', '10:30'],
		['C4', '=1/0'],
	];
	session.batch('Fill', () => {
		for (const [ref, input] of inputs) {
			const at = parseAddress(ref)!;
			session.setCellInput(0, at.row, at.col, input);
		}
		session.setCellInput(1, 1, 1, '0.08');
	});
	session.createTable(0, parseRange('A1:D4')!, true);
	put(workbook, 0, 'F6', 1.5, '[h]:mm');
	put(workbook, 0, 'E5', 45385.25, 'd-mmm');
	workbook.sheets[1]!.state = 'hidden';
	workbook.activeSheet = 0;
	workbook.definedNames.push(
		{ name: 'TaxRate', formula: 'Lookup!$B$2' },
		{ name: 'Origin', formula: 'Sales!$A$1', localSheet: 0 },
		{ name: 'Secret', formula: 'Lookup!$B$2', hidden: true },
	);
	return saveXlsx(workbook);
}

describe('inspectXlsx', () => {
	it('describes each sheet: used range, size, visibility and tables, plus the defined names', async () => {
		const result = await inspectXlsx(await sample());
		expect(result.activeSheet).toBe(0);
		expect(result.sheets).toEqual([
			{
				index: 0,
				name: 'Sales',
				state: 'visible',
				usedRange: 'A1:F6',
				rows: 6,
				columns: 6,
				tables: [{ name: 'Table1', range: 'A1:D4', headerRow: true, totalsRow: false }],
			},
			{
				index: 1,
				name: 'Lookup',
				state: 'hidden',
				usedRange: 'B2',
				rows: 1,
				columns: 1,
				tables: [],
			},
			{
				index: 2,
				name: 'Empty',
				state: 'visible',
				usedRange: null,
				rows: 0,
				columns: 0,
				tables: [],
			},
		]);
		// Hidden names (Excel's Name Manager does not list them) are left out.
		expect(result.definedNames).toEqual([
			{ name: 'TaxRate', refersTo: 'Lookup!$B$2' },
			{ name: 'Origin', refersTo: 'Sales!$A$1', sheetIndex: 0 },
		]);
		expect(result).not.toHaveProperty('definedNamesOmitted');
	});

	it('bounds the defined-name list and says how many it left out', async () => {
		const workbook = createWorkbook();
		for (let i = 0; i < 520; i++)
			workbook.definedNames.push({ name: `Name${i}`, formula: `Sheet1!$A$${i + 1}` });
		const result = await inspectXlsx(await saveXlsx(workbook));
		expect(result.definedNames).toHaveLength(500);
		expect(result.definedNamesOmitted).toBe(20);
	});
});

describe('readXlsxRange', () => {
	it('returns the displayed text, the number format and an ISO form of dates next to each value', async () => {
		const { cells } = await readXlsxRange(await sample(), 0, 'A1:F6');
		const at = (address: string) => cells.find((cell) => cell.address === address);
		expect(at('A2')).toEqual({ address: 'A2', value: 'North', text: 'North' });
		expect(at('B2')).toEqual({
			address: 'B2',
			value: 45385,
			text: '4/3/2024',
			numberFormat: 'm/d/yyyy',
			iso: '2024-04-03',
		});
		expect(at('B3')).toMatchObject({ text: '2024-04-03 10:30', iso: '2024-04-03T10:30:00' });
		expect(at('B4')).toMatchObject({ text: '10:30', numberFormat: 'h:mm', iso: '10:30:00' });
		expect(at('C2')).toMatchObject({ value: 0.153, text: '15.30%', numberFormat: '0.00%' });
		expect(at('C2')).not.toHaveProperty('iso');
		expect(at('D2')).toMatchObject({ value: 1234.5, text: '$1,234.50 ' });
		expect(at('D3')).toEqual({ address: 'D3', value: 2469, formula: 'D2*2', text: '2469' });
		expect(at('A4')).toEqual({ address: 'A4', value: true, text: 'TRUE' });
		expect(at('C4')).toMatchObject({ formula: '1/0', text: '#DIV/0!' });
		// The format hides the time; the ISO form keeps the value.
		expect(at('E5')).toMatchObject({ text: '3-Apr', iso: '2024-04-03T06:00:00' });
		// A duration is not a point in time.
		expect(at('F6')).toMatchObject({ value: 1.5, text: '36:00', numberFormat: '[h]:mm' });
		expect(at('F6')).not.toHaveProperty('iso');
		expect(at('E1')).toEqual({ address: 'E1', value: null, text: '' });
	});
});

/** One column of values and formats, read back as ISO forms (`undefined` where none is given). */
async function isoForms(date1904: boolean, cells: [number, string][]) {
	const workbook = createWorkbook();
	workbook.date1904 = date1904;
	cells.forEach(([value, numFmt], row) => put(workbook, 0, `A${row + 1}`, value, numFmt));
	const read = await readXlsxRange(await saveXlsx(workbook), 0, `A1:A${cells.length}`);
	return read.cells.map((cell) => cell.iso);
}

describe('readXlsxRange ISO dates', () => {
	it('keeps fractions of a second and follows the format section that shows the value', async () => {
		expect(
			await isoForms(false, [
				[0.5 + 0.5 / 86_400, 'hh:mm:ss.000'],
				[45385.5 + 0.25 / 86_400, 'yyyy-mm-dd hh:mm:ss.000'],
				[45385, 'yyyy-mm-dd;[h]:mm'],
				[5, '0.00;yyyy'],
				[2958465.5, 'yyyy-mm-dd hh:mm'],
				[2958466, 'yyyy-mm-dd'],
				// Far beyond what a JavaScript date can hold: no ISO form, and the read still succeeds.
				[1_000_000_000, 'yyyy-mm-dd'],
			]),
		).toEqual([
			'12:00:00.500',
			'2024-04-03T12:00:00.250',
			'2024-04-03',
			undefined,
			'9999-12-31T12:00:00',
			undefined,
			undefined,
		]);
	});

	it('reads the 1904 date system, whose day 0 is a real date', async () => {
		expect(
			await isoForms(true, [
				[0, 'yyyy-mm-dd'],
				[0.25, 'yyyy-mm-dd hh:mm'],
				[0.25, 'h:mm'],
				[2957003, 'yyyy-mm-dd'],
				[2957004, 'yyyy-mm-dd'],
			]),
		).toEqual(['1904-01-01', '1904-01-01T06:00:00', '06:00:00', '9999-12-31', undefined]);
	});
});

describe('findXlsxCells', () => {
	it('finds values as displayed and formulas as typed, with Excel wildcards', async () => {
		const bytes = await sample();
		const north = await findXlsxCells(bytes, 'nor*', { lookIn: 'values' });
		expect(north.matches).toEqual([
			{ sheetIndex: 0, sheet: 'Sales', address: 'A2', value: 'North', text: 'North' },
		]);
		expect(north.truncated).toBe(false);
		// Values are searched as displayed: the date reads 4/3/2024, not 45385.
		expect((await findXlsxCells(bytes, '4/3/2024', { lookIn: 'values' })).matches).toMatchObject([
			{ address: 'B2', iso: '2024-04-03' },
		]);
		expect((await findXlsxCells(bytes, '45385', { lookIn: 'values' })).matches).toEqual([]);
		// Formulas (Excel's default) search what was typed.
		expect((await findXlsxCells(bytes, 'D2*2', { wildcards: false })).matches).toMatchObject([
			{ address: 'D3', formula: 'D2*2', value: 2469 },
		]);
	});

	it('honours case, whole-cell, sheet and range limits', async () => {
		const bytes = await sample();
		const addresses = async (query: string, options: Parameters<typeof findXlsxCells>[2]) =>
			(await findXlsxCells(bytes, query, options)).matches.map(
				(m) => `${m.sheetIndex}!${m.address}`,
			);
		expect(await addresses('north', { matchCase: true })).toEqual([]);
		expect(await addresses('o', { wholeCell: true })).toEqual([]);
		expect(await addresses('o', { lookIn: 'values' })).toEqual([
			'0!A1',
			'0!B1',
			'0!D1',
			'0!A2',
			'0!A3',
		]);
		expect(await addresses('o', { lookIn: 'values', sheetIndex: 0, range: 'A2:B3' })).toEqual([
			'0!A2',
			'0!A3',
		]);
		expect(await addresses('0.08', {})).toEqual(['1!B2']);
		await expect(findXlsxCells(bytes, 'x', { sheetIndex: 9 })).rejects.toThrow('Invalid sheet');
		await expect(findXlsxCells(bytes, 'x', { range: 'A1:B2' })).rejects.toThrow('sheetIndex');
		await expect(findXlsxCells(bytes, '', {})).rejects.toThrow('empty');
	});

	it('bounds the matches and counts the rest', async () => {
		const result = await findXlsxCells(await sample(), '*', { maxResults: 3 });
		expect(result.matches).toHaveLength(3);
		expect(result.truncated).toBe(true);
		expect(result.total).toBeGreaterThan(3);
		await expect(findXlsxCells(await sample(), '*', { maxResults: 1001 })).rejects.toThrow('1000');
	});
});

describe('readXlsxTable', () => {
	it('reads a table by name, ignoring case, with its column names and cells', async () => {
		const result = await readXlsxTable(await sample(), 'table1');
		expect(result).toMatchObject({
			sheetIndex: 0,
			sheet: 'Sales',
			name: 'Table1',
			range: 'A1:D4',
			headerRow: true,
			totalsRow: false,
			columns: ['Region', 'Opened', 'Share', 'Amount'],
		});
		expect(result.cells).toHaveLength(16);
		expect(result.cells[4]).toEqual({ address: 'A2', value: 'North', text: 'North' });
		expect(result.cells[5]).toMatchObject({ address: 'B2', iso: '2024-04-03' });
		await expect(readXlsxTable(await sample(), 'Missing')).rejects.toThrow('No table named');
	});
});
