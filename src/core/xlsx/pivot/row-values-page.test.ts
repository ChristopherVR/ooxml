import { readdirSync, readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { children, elements, first, NS, parseXml } from '../../xml/index';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import type { PivotTable, PivotTableOptions, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { styleAt } from '../styles';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';

const dir = new URL('../__fixtures__/', import.meta.url);
const fixture = (name: string) => readFileSync(new URL(name, dir));
const PAGE = readdirSync(dir).filter((name) => /^excel-pivot-rvp-.*\.xlsx$/.test(name));

/** The options that author the pivot a file holds, so the writer can be compared with Excel's. */
function optionsOf(pivot: PivotTable): PivotTableOptions {
	return {
		source: pivot.source,
		rowField: pivot.rowField,
		layout: pivot.layout!,
		subtotalPosition: pivot.subtotalPosition!,
		...(pivot.secondRowField ? { secondRowField: pivot.secondRowField } : {}),
		...(pivot.moreRowFields ? { moreRowFields: pivot.moreRowFields } : {}),
		...(pivot.subtotals !== undefined ? { subtotals: pivot.subtotals } : {}),
		...(pivot.repeatLabels !== undefined ? { repeatLabels: pivot.repeatLabels } : {}),
		...(pivot.columnField ? { columnField: pivot.columnField } : {}),
		...(pivot.pageFilter ? { pageFilter: pivot.pageFilter } : {}),
		valueField: pivot.valueField,
		aggregate: pivot.aggregate,
		additionalValues: pivot.additionalValues!,
		valuesAxis: 'rows',
		...(typeof pivot.valuesPosition !== 'undefined'
			? { valuesPosition: pivot.valuesPosition }
			: {}),
	};
}

const summaryOf = (wb: Workbook) => wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
const indentsOf = (wb: Workbook, sheet: number, pivot: PivotTable, rows: number) =>
	Array.from(
		{ length: rows },
		(_v, r) =>
			styleAt(
				wb,
				getCell(wb.sheets[sheet]!, pivot.location.start.row + r, pivot.location.start.col)?.styleId,
			).alignment?.indent ?? 0,
	);

describe('Values on rows with a report filter', () => {
	it('has real Excel fixtures to compare with', () =>
		expect(PAGE.length).toBeGreaterThanOrEqual(6));

	for (const file of PAGE) {
		it(`reads, rewrites and reauthors ${file} the way Excel wrote it`, async () => {
			const bytes = fixture(file);
			const wb = await loadXlsx(bytes);
			expect(wb.warnings).toEqual([]);
			const s = summaryOf(wb);
			const pivot = wb.sheets[s]!.pivotTables![0]!;
			expect(pivot.valuesAxis).toBe('rows');
			expect(pivot.pageFilter).toBeDefined();
			const grid = pivotGrid(pivot);
			// Excel's cached cells are what the grid reproduces, indentation included.
			expect(indentsOf(wb, s, pivot, grid.values.length)).toEqual(
				grid.values.map((_row, r) => grid.indents?.[r] ?? 0),
			);

			// An untouched load saves the pivot parts byte for byte.
			const original = await JSZip.loadAsync(bytes);
			const saved = await JSZip.loadAsync(await saveXlsx(wb));
			for (const part of Object.keys(original.files).filter((p) => p.startsWith('xl/pivot')))
				expect(await saved.file(part)?.async('string')).toBe(
					await original.file(part)!.async('string'),
				);

			// A refresh keeps the cells and the report filter through save and reopen.
			const edit = createEditSession(wb);
			edit.refreshPivotTable(s, pivot.name);
			const reopened = await loadXlsx(await saveXlsx(wb));
			expect(reopened.warnings).toEqual([]);
			const again = reopened.sheets[s]!.pivotTables![0]!;
			expect(again.pageFilter).toEqual(pivot.pageFilter);
			expect(pivotGrid(again).values).toEqual(grid.values);

			// A pivot authored from the same options has Excel's structure.
			const name = edit.createPivotTable(s, { row: 0, col: 12 }, optionsOf(pivot));
			const authored = wb.sheets[s]!.pivotTables!.find((entry) => entry.name === name)!;
			// The fixture renames its first value field to Data; a new pivot names it by default.
			expect(pivotGrid(authored).values).toEqual(
				grid.values.map((row) =>
					row.map((v) => (typeof v === 'string' ? v.replaceAll('Data', 'Sum of Amount') : v)),
				),
			);
			const zip = await JSZip.loadAsync(await saveXlsx(wb));
			const parts = Object.keys(zip.files).filter((p) =>
				/^xl\/pivotTables\/pivotTable\d+\.xml$/.test(p),
			);
			const roots = await Promise.all(
				parts.map(async (p) => parseXml(await zip.file(p)!.async('string')).documentElement),
			);
			const [excel, ours] = ['RowValuesPivot', name].map((n) =>
				roots.find((root) => root.getAttribute('name') === n)!,
			);
			const shape = (root: Element) => ({
				location: Object.fromEntries(
					['firstHeaderRow', 'firstDataRow', 'firstDataCol', 'rowPageCount', 'colPageCount'].map(
						(k) => [k, first(root, 'location', NS.x)!.getAttribute(k)],
					),
				),
				rowFields: elements(first(root, 'rowFields', NS.x)!).map((f) => f.getAttribute('x')),
				colFields: elements(
					first(root, 'colFields', NS.x) ?? root.ownerDocument.createElement('x'),
				).map((f) => f.getAttribute('x')),
				pageFields: elements(first(root, 'pageFields', NS.x)!).map((f) =>
					['fld', 'item', 'hier'].map((k) => f.getAttribute(k)),
				),
				rowItems: children(first(root, 'rowItems', NS.x)!, 'i', NS.x).map((i) => [
					i.getAttribute('t'),
					i.getAttribute('r'),
					i.getAttribute('i'),
					elements(i).map((x) => x.getAttribute('v')),
				]),
				colItems: children(first(root, 'colItems', NS.x)!, 'i', NS.x).map((i) => [
					i.getAttribute('t'),
					i.getAttribute('r'),
					i.getAttribute('i'),
					elements(i).map((x) => x.getAttribute('v')),
				]),
				dataOnRows: root.getAttribute('dataOnRows'),
				dataPosition: root.getAttribute('dataPosition'),
			});
			expect(shape(ours!)).toEqual(shape(excel!));
		});
	}

	it('moves a Values-on-columns report filter to rows and back with undo', async () => {
		const wb = await loadXlsx(fixture('excel-pivot-rvp-last-compact.xlsx'));
		const s = summaryOf(wb);
		const edit = createEditSession(wb);
		const name = wb.sheets[s]!.pivotTables![0]!.name;
		const before = pivotGrid(wb.sheets[s]!.pivotTables![0]!).values;
		edit.refreshPivotTable(s, name, { valuesAxis: 'columns' });
		const columns = wb.sheets[s]!.pivotTables![0]!;
		expect(columns.pageFilter).toEqual({ field: 'Year' });
		expect(pivotGrid(columns).values[2]).toEqual(['Row Labels', 'Data', 'Count of Amount']);
		edit.refreshPivotTable(s, name, { valuesAxis: 'rows' });
		expect(pivotGrid(wb.sheets[s]!.pivotTables![0]!).values).toEqual(before);
		edit.undo();
		edit.undo();
		expect(pivotGrid(wb.sheets[s]!.pivotTables![0]!).values).toEqual(before);
	});

	it('still refuses a report filter on a field that is on an axis', async () => {
		const wb = await loadXlsx(fixture('excel-pivot-rvp-last-compact.xlsx'));
		const s = summaryOf(wb);
		const edit = createEditSession(wb);
		expect(() =>
			edit.refreshPivotTable(s, wb.sheets[s]!.pivotTables![0]!.name, {
				pageFilter: { field: 'Category' },
			}),
		).toThrow('separate source field');
	});
});

describe('compact indentation of Values on rows', () => {
	const COMPACT = readdirSync(dir).filter(
		(name) => /^excel-pivot-.*compact.*\.xlsx$/.test(name) && !name.startsWith('excel-pivot-rvp-'),
	);
	for (const file of COMPACT) {
		it(`matches the indents Excel wrote in ${file}`, async () => {
			const wb = await loadXlsx(fixture(file));
			const s = summaryOf(wb);
			if (s < 0) return;
			const pivot = wb.sheets[s]!.pivotTables![0]!;
			if (pivot.valuesAxis !== 'rows') return;
			const grid = pivotGrid(pivot);
			expect(indentsOf(wb, s, pivot, grid.values.length)).toEqual(
				grid.values.map((_row, r) => grid.indents?.[r] ?? 0),
			);
		});
	}
});
