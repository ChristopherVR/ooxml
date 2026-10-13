import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { getCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';

const fixture = readFileSync(new URL('../__fixtures__/excel-pivot.xlsx', import.meta.url));

for (const suffix of ['', '-no-subtotals', '-repeat'])
	it(`reads and refreshes native two-row layout${suffix}`, async () => {
		const bytes = readFileSync(
			new URL(`../__fixtures__/excel-pivot-rows${suffix}.xlsx`, import.meta.url),
		);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const summary = wb.sheets.findIndex((sheet) => sheet.name === 'Summary');
		const data = wb.sheets.findIndex((sheet) => sheet.name === 'Data');
		const pivot = wb.sheets[summary]!.pivotTables![0]!;
		expect(pivot).toMatchObject({
			secondRowField: 'Region',
			subtotals: suffix !== '-no-subtotals',
			repeatLabels: suffix === '-repeat',
		});
		const original = await JSZip.loadAsync(bytes);
		const unchanged = await JSZip.loadAsync(await saveXlsx(wb));
		for (const path of Object.keys(original.files).filter((path) => path.startsWith('xl/pivot')))
			expect(await unchanged.file(path)?.async('string')).toBe(
				await original.file(path)!.async('string'),
			);
		const edit = createEditSession(wb);
		edit.setCellInput(data, 1, 2, '100');
		edit.refreshPivotTable(summary, pivot.name);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(getCell(reloaded.sheets[summary]!, 1, 2)?.value).toBe(107);
	});

it('preserves unsupported custom subtotals and field extensions without admission', async () => {
	const bytes = readFileSync(
		new URL('../__fixtures__/excel-pivot-rows-repeat.xlsx', import.meta.url),
	);
	const path = 'xl/pivotTables/pivotTable1.xml';
	for (const patch of [
		(xml: string) => xml.replace('<item t="default"/>', '<item t="sum"/>'),
		(xml: string) => xml.replace('fillDownLabels="1"', 'fillDownLabels="0"'),
		(xml: string) => xml.replace('fillDownLabels="1"', 'fillDownLabels="1" insertBlankRow="1"'),
	]) {
		const zip = await JSZip.loadAsync(bytes);
		const xml = patch(await zip.file(path)!.async('string'));
		zip.file(path, xml);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(wb.warnings.some((warning) => warning.includes('preserved'))).toBe(true);
		expect(wb.sheets.find((sheet) => sheet.name === 'Summary')?.pivotTables).toBeUndefined();
		const output = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await output.file(path)!.async('string')).toBe(xml);
	}
});

it('reads a native Excel PivotTable, preserves untouched parts, and refreshes its cache', async () => {
	const book = await loadXlsx(fixture);
	expect(book.warnings).toEqual([]);
	const summary = book.sheets.findIndex((sheet) => sheet.name === 'Summary');
	const data = book.sheets.findIndex((sheet) => sheet.name === 'Data');
	const pivot = book.sheets[summary]!.pivotTables![0]!;
	expect(pivot).toMatchObject({ name: 'SalesPivot', source: 'Data!A1:C6', aggregate: 'sum' });
	const before = await JSZip.loadAsync(fixture);
	const untouched = await JSZip.loadAsync(await saveXlsx(book));
	for (const path of Object.keys(before.files).filter((path) => path.startsWith('xl/pivot')))
		expect(await untouched.file(path)?.async('string')).toBe(
			await before.file(path)!.async('string'),
		);
	const edit = createEditSession(book);
	edit.setCellInput(data, 1, 2, '100');
	edit.refreshPivotTable(summary, pivot.name);
	const refreshed = await loadXlsx(await saveXlsx(book));
	expect(refreshed.warnings).toEqual([]);
	expect(getCell(refreshed.sheets[summary]!, 4, 3)?.value).toBe(147);
	edit.deletePivotTable(summary, pivot.name);
	const deleted = await JSZip.loadAsync(await saveXlsx(book));
	expect(Object.keys(deleted.files).filter((path) => path.startsWith('xl/pivot'))).toEqual([]);
});

it('preserves unsupported native layouts without pretending they can be authored', async () => {
	const zip = await JSZip.loadAsync(fixture);
	const path = 'xl/pivotTables/pivotTable1.xml';
	const xml = (await zip.file(path)!.async('string')).replace('compact="0"', 'compact="1"');
	zip.file(path, xml);
	const book = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(book.sheets.every((sheet) => !sheet.pivotTables)).toBe(true);
	expect(book.warnings.some((warning) => warning.includes('cannot be refreshed'))).toBe(true);
	const saved = await JSZip.loadAsync(await saveXlsx(book));
	expect(await saved.file(path)!.async('string')).toBe(xml);
	for (const cache of Object.keys(zip.files).filter((name) => name.startsWith('xl/pivotCache')))
		expect(await saved.file(cache)?.async('string')).toBe(await zip.file(cache)!.async('string'));
});

it.each([
	[
		'table extension',
		'xl/pivotTables/pivotTable1.xml',
		'</extLst>',
		'<ext uri="unknown"><extra xmlns="urn:vendor"/></ext></extLst>',
	],
	[
		'cache extension',
		'xl/pivotCache/pivotCacheDefinition1.xml',
		'</extLst>',
		'<ext uri="unknown"><extra xmlns="urn:vendor"/></ext></extLst>',
	],
	['field formatting', 'xl/pivotCache/pivotCacheDefinition1.xml', 'numFmtId="0"', 'numFmtId="14"'],
	['record metadata', 'xl/pivotCache/pivotCacheRecords1.xml', '<r>', '<r vendor="keep">'],
	['field setting', 'xl/pivotTables/pivotTable1.xml', 'outline="0"', 'outline="1"'],
])(
	'preserves unsupported %s and reserves its table name',
	async (_label, part, find, replacement) => {
		const zip = await JSZip.loadAsync(fixture);
		const xml = (await zip.file(part)!.async('string')).replace(find, replacement);
		zip.file(part, xml);
		const book = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(book.sheets[0]!.pivotTables).toBeUndefined();
		const edit = createEditSession(book);
		expect(() =>
			edit.createPivotTable(
				0,
				{ row: 8, col: 0 },
				{ source: 'Data!A1:C6', rowField: 'Category', valueField: 'Amount', name: 'salespivot' },
			),
		).toThrow('unique');
		edit.createPivotTable(
			0,
			{ row: 8, col: 0 },
			{ source: 'Data!A1:C6', rowField: 'Category', valueField: 'Amount' },
		);
		const saved = await JSZip.loadAsync(await saveXlsx(book));
		expect(await saved.file(part)!.async('string')).toBe(xml);
		expect((await loadXlsx(await saveXlsx(book))).sheets[0]!.pivotTables).toHaveLength(1);
	},
);

it('reserves default names and preserves tables using a shared native cache', async () => {
	const zip = await JSZip.loadAsync(fixture);
	const tablePath = 'xl/pivotTables/pivotTable1.xml';
	const first = (await zip.file(tablePath)!.async('string')).replace(
		'name="SalesPivot"',
		'name="PivotTable1"',
	);
	zip.file(tablePath, first);
	zip.file(
		'xl/pivotTables/pivotTable2.xml',
		first.replace('name="PivotTable1"', 'name="PivotTable2"'),
	);
	zip.file(
		'xl/pivotTables/_rels/pivotTable2.xml.rels',
		await zip.file('xl/pivotTables/_rels/pivotTable1.xml.rels')!.async('string'),
	);
	const relPath = 'xl/worksheets/_rels/sheet1.xml.rels';
	zip.file(
		relPath,
		(await zip.file(relPath)!.async('string')).replace(
			'</Relationships>',
			'<Relationship Id="rId77" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/></Relationships>',
		),
	);
	const book = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(book.sheets[0]!.pivotTables).toBeUndefined();
	const name = createEditSession(book).createPivotTable(
		0,
		{ row: 8, col: 0 },
		{ source: 'Data!A1:C6', rowField: 'Category', valueField: 'Amount' },
	);
	expect(name).toBe('PivotTable3');
	const saved = await JSZip.loadAsync(await saveXlsx(book));
	expect(await saved.file(tablePath)!.async('string')).toBe(first);
	expect(await saved.file('xl/pivotTables/pivotTable2.xml')!.async('string')).toBe(
		first.replace('name="PivotTable1"', 'name="PivotTable2"'),
	);
});
