import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../../xml/index';
import { applyChartSpaceDataSheetEdits, chartSpaceDataSheet } from '../edit/data-sheet';
import { DEFAULT_CHART_DATA_SHEET, newChartSpace } from '../edit/new-chart';
import { createChartWorkbook } from './create';
import { standardChartLayoutSheet } from './layout';
import { chartSpaceWorkbookWrites } from './space-writes';
import { writeChartWorkbookUpdates } from './writer';

/** `ref -> value` of every cell of the first worksheet. */
async function cells(bytes: Uint8Array): Promise<Record<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	const doc = parseXml(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
	const out: Record<string, string> = {};
	for (const cell of Array.from(doc.getElementsByTagName('c'))) {
		const v = cell.getElementsByTagName('v')[0] ?? cell.getElementsByTagName('t')[0];
		out[cell.getAttribute('r')!] = v?.textContent ?? '';
	}
	return out;
}

describe('the embedded workbook of a new chart', () => {
	it('holds the data sheet in the standard layout the chart formulas name', async () => {
		const space = newChartSpace();
		const bytes = await createChartWorkbook(DEFAULT_CHART_DATA_SHEET);
		const grid = await cells(bytes);
		expect(grid).toMatchObject({ B1: 'Series 1', D1: 'Series 3', A2: 'Category 1', B2: '4.3' });
		expect(grid.D5).toBe('5');
		expect(standardChartLayoutSheet(space, 4)).toBe('Sheet1');
		// The workbook already agrees with the cache: nothing to write.
		const result = await writeChartWorkbookUpdates(bytes, chartSpaceWorkbookWrites(space));
		expect(result).toEqual({ unresolved: 0 });
	});

	it('follows value, name and structural edits cell for cell', async () => {
		const bytes = await createChartWorkbook(DEFAULT_CHART_DATA_SHEET);
		const edited = applyChartSpaceDataSheetEdits(newChartSpace(), [
			{ kind: 'value', series: 0, category: 1, value: 9.5 },
			{ kind: 'series-name', series: 1, name: 'Profit' },
			{ kind: 'category-label', category: 0, label: 'North' },
			{ kind: 'add-series' },
			{ kind: 'add-category' },
		]);
		expect(chartSpaceDataSheet(edited).series.map((series) => series.name)).toEqual([
			'Series 1',
			'Profit',
			'Series 3',
			'Series 4',
		]);
		expect(standardChartLayoutSheet(edited, 5)).toBe('Sheet1');
		const result = await writeChartWorkbookUpdates(bytes, chartSpaceWorkbookWrites(edited));
		expect(result.unresolved).toBe(0);
		const grid = await cells(result.bytes!);
		expect(grid).toMatchObject({
			B3: '9.5',
			C1: 'Profit',
			A2: 'North',
			E1: 'Series 4',
			A6: 'Cat 5',
			B6: '0',
			E6: '0',
		});
	});

	it('keeps a structural edit cache-only when the chart reads a non-standard range', async () => {
		const bytes = await createChartWorkbook(DEFAULT_CHART_DATA_SHEET);
		const space = newChartSpace();
		const moved = structuredClone(space);
		moved.plotArea.groups[0]!.series[0]!.values!.formula = 'Sheet1!$B$2:$B$4';
		const edited = applyChartSpaceDataSheetEdits(moved, [{ kind: 'add-category' }]);
		const result = await writeChartWorkbookUpdates(bytes, chartSpaceWorkbookWrites(edited));
		expect(result.unresolved).toBeGreaterThan(0);
	});

	it('drops a formula and empties a blank point when writing over a cell', async () => {
		const zip = await JSZip.loadAsync(await createChartWorkbook(DEFAULT_CHART_DATA_SHEET));
		const path = 'xl/worksheets/sheet1.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replace(
				'<c r="B2"><v>4.3</v></c>',
				'<c r="B2" s="3"><f>1+1</f><v>2</v></c>',
			),
		);
		const result = await writeChartWorkbookUpdates(
			await zip.generateAsync({ type: 'uint8array' }),
			[{ formula: 'Sheet1!$B$2:$B$3', isNumeric: true, values: ['7', ''] }],
		);
		const xml = await (await JSZip.loadAsync(result.bytes!)).file(path)!.async('string');
		expect(xml).toContain('<c r="B2" s="3"><v>7</v></c>');
		expect(xml).toContain('<c r="B3"/>');
	});
});

describe('the data table Word and PowerPoint put behind a chart', () => {
	/** The sample workbook with Word's `Table1` over A1:D5 and its names in shared strings. */
	async function withTable(): Promise<Uint8Array> {
		const zip = await JSZip.loadAsync(await createChartWorkbook(DEFAULT_CHART_DATA_SHEET));
		const sheet = 'xl/worksheets/sheet1.xml';
		zip.file(
			sheet,
			(await zip.file(sheet)!.async('string'))
				.replace(
					'<c r="B1" t="inlineStr"><is><t xml:space="preserve">Series 1</t></is></c>',
					'<c r="B1" t="s"><v>0</v></c>',
				)
				.replace('<row r="1">', '<row r="1" spans="1:4">')
				.replace(
					'</worksheet>',
					'<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
				),
		);
		zip.file(
			'xl/sharedStrings.xml',
			'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Series 1</t></si></sst>',
		);
		zip.file(
			'xl/worksheets/_rels/sheet1.xml.rels',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
		);
		zip.file(
			'xl/tables/table1.xml',
			'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:D5"><tableColumns count="4"><tableColumn id="1" name=" "/><tableColumn id="2" name="Series 1"/><tableColumn id="3" name="Series 2"/><tableColumn id="4" name="Series 3"/></tableColumns></table>',
		);
		return zip.generateAsync({ type: 'uint8array' });
	}

	it('renames and grows the table with the data and keeps unchanged shared strings', async () => {
		const edited = applyChartSpaceDataSheetEdits(newChartSpace(), [
			{ kind: 'series-name', series: 1, name: 'Renamed' },
			{ kind: 'add-series' },
			{ kind: 'add-category' },
		]);
		const result = await writeChartWorkbookUpdates(
			await withTable(),
			chartSpaceWorkbookWrites(edited),
		);
		expect(result.unresolved).toBe(0);
		const zip = await JSZip.loadAsync(result.bytes!);
		const table = await zip.file('xl/tables/table1.xml')!.async('string');
		expect(table).toContain('ref="A1:E6"');
		expect(table).toContain('<tableColumns count="5">');
		expect(table).toContain('name="Renamed"');
		expect(table).toContain('<tableColumn id="1" name=" "/>');
		expect(table).toContain('<tableColumn id="5" name="Series 4"/>');
		const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(sheet).toContain('<c r="B1" t="s"><v>0</v></c>');
		expect(sheet).toContain('<dimension ref="A1:E6"/>');
		expect(sheet).not.toContain('spans=');
	});
});
