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
