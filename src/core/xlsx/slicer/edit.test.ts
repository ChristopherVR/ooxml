import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { selectionAfterClick } from './select';
import { slicerItems } from './items';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-slicer-${name}.xlsx`, import.meta.url));
const text = async (bytes: Uint8Array, name: string) =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';

it('lists table slicer items with their selected and no-data states', async () => {
	const wb = await loadXlsx(fixture('table'));
	const items = (name: string) =>
		slicerItems(
			wb,
			wb.slicers!.caches.find((c) => c.name === name)!,
		).map((i) => [i.label, i.state]);
	expect(items('Slicer_Category')).toEqual([
		['Fruit', 'selected'],
		['Nuts', 'selected'],
		['Veg', 'unselected'],
	]);
	// Category hides Veg, so Region items still have data from the visible rows only.
	expect(items('Slicer_Region')).toEqual([
		['East', 'selected'],
		['West', 'selected'],
	]);
});

it('clicking a table slicer item filters rows, saves the filter and reloads it', async () => {
	const wb = await loadXlsx(fixture('table'));
	const edit = createEditSession(wb);
	edit.setSlicerSelection('Slicer_Category', ['Veg']);
	const sheet = wb.sheets[0]!;
	expect(sheet.rowInfo.get(1)?.hidden).toBe(true);
	expect(sheet.rowInfo.get(3)?.hidden).toBeUndefined();
	const saved = await saveXlsx(wb);
	const table = await text(saved, 'xl/tables/table1.xml');
	expect(table).toContain('<filter val="Veg"/>');
	expect(table).not.toContain('val="Fruit"');
	const again = await loadXlsx(saved);
	expect(again.sheets[0]!.tables[0]!.columns[0]!.filter).toEqual({ values: ['Veg'] });
	// The slicer parts are untouched by a table selection.
	expect(await text(saved, 'xl/slicerCaches/slicerCache1.xml')).toBe(
		await text(new Uint8Array(fixture('table')), 'xl/slicerCaches/slicerCache1.xml'),
	);
	edit.undo();
	expect(sheet.rowInfo.get(3)?.hidden).toBe(true);
	edit.setSlicerSelection('Slicer_Category', null);
	expect(sheet.rowInfo.get(1)?.hidden).toBeUndefined();
	expect(sheet.tables[0]!.columns[0]!.filter).toBeUndefined();
});

it('filters a PivotTable through its slicer, saves hidden items and the cache selection', async () => {
	const wb = await loadXlsx(fixture('pivot'));
	const edit = createEditSession(wb);
	const keyOf = (label: string) =>
		slicerItems(wb, wb.slicers!.caches[0]!).find((item) => item.label === label)!.key;
	edit.setSlicerSelection('Slicer_Category', [keyOf('Nuts'), keyOf('Veg')]);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.itemFilters).toEqual([{ field: 'Category', hidden: ['Fruit'] }]);
	const sheet = wb.sheets[0]!;
	const column = [2, 3, 4, 5].map((row) => getCell(sheet, row, 0)?.value);
	expect(column).toContain('Nuts');
	expect(column).not.toContain('Fruit');
	const saved = await saveXlsx(wb);
	expect(await text(saved, 'xl/pivotTables/pivotTable1.xml')).toMatch(/<item h="1" x="0"\/>/);
	const cache = await text(saved, 'xl/slicerCaches/slicerCache1.xml');
	expect(cache).toMatch(/<i x="0" nd?="?1?"?\/?>|<i x="0"\/>|<i x="0"/);
	expect(cache).toContain(`pivotCacheId="${pivot.slicerCacheId}"`);
	const again = await loadXlsx(saved);
	expect(again.warnings).toEqual([]);
	expect(again.sheets[0]!.pivotTables![0]!.itemFilters).toEqual(pivot.itemFilters);
	expect(
		slicerItems(again, again.slicers!.caches[0]!).map((item) => [item.label, item.selected]),
	).toEqual([
		['Fruit', false],
		['Nuts', true],
		['Veg', true],
	]);
});

async function created() {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
	const edit = createEditSession(wb);
	const rows = [
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', '10'],
		['Fruit', 'West', '20'],
		['Veg', 'East', '5'],
		['Veg', 'West', '15'],
	];
	rows.forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
	edit.createTable(0, { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } }, true);
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{ source: 'Data!A1:C5', rowField: 'Category', valueField: 'Amount' },
	);
	return { wb, edit };
}

it('inserts table and PivotTable slicers that save, reload and filter', async () => {
	const { wb, edit } = await created();
	const tableSlicer = edit.insertSlicer(0, { kind: 'table', table: 'Table1', column: 'Region' });
	const pivotSlicer = edit.insertSlicer(1, {
		kind: 'pivot',
		pivot: 'PivotTable1',
		field: 'Category',
	});
	expect([tableSlicer, pivotSlicer]).toEqual(['Region', 'Category']);
	expect(wb.sheets[0]!.drawings.map((d) => d.kind)).toEqual(['slicer']);
	const saved = await saveXlsx(wb);
	const again = await loadXlsx(saved);
	expect(again.warnings).toEqual([]);
	expect(again.slicers?.slicers.map((s) => s.name).sort()).toEqual(['Category', 'Region']);
	expect(again.sheets[1]!.pivotTables![0]!.slicerCacheId).toBeTypeOf('number');
	const reopened = createEditSession(again);
	reopened.setSlicerSelection('Slicer_Region', ['West']);
	expect(again.sheets[0]!.rowInfo.get(1)?.hidden).toBe(true);
	reopened.undo();
	expect(again.sheets[0]!.rowInfo.get(1)?.hidden).toBeUndefined();
	reopened.deleteSlicer('Region');
	expect(again.slicers?.slicers.map((s) => s.name)).toEqual(['Category']);
	expect(again.definedNames.map((n) => n.name)).toEqual(['Slicer_Category']);
	const final = await saveXlsx(again);
	const names = Object.keys((await JSZip.loadAsync(final)).files).filter((n) =>
		/slicer(s|Caches)\//i.test(n),
	);
	expect(names).toHaveLength(2);
	expect(names.filter((n) => n.includes('slicerCaches'))).toHaveLength(1);
});

it('works on workbooks built without optional model parts', async () => {
	const { wb, edit } = await created();
	const views = wb.sheets.map((sheet) => sheet.view);
	for (const sheet of wb.sheets) delete (sheet as { view?: unknown }).view;
	expect(wb.slicers).toBeUndefined();
	expect(() => edit.setSlicerSelection('missing', null)).toThrow();
	edit.insertSlicer(0, { kind: 'table', table: 'Table1', column: 'Region' });
	edit.deletePivotTable(1, 'PivotTable1');
	wb.sheets.forEach((sheet, i) => (sheet.view = views[i]!));
	expect((await loadXlsx(await saveXlsx(wb))).slicers?.slicers).toHaveLength(1);
});

it('selects like Excel: plain click, toggle, range and emptying the selection', () => {
	const item = (key: string, selected: boolean) => ({
		key,
		label: key,
		selected,
		hasData: true,
		state: selected ? ('selected' as const) : ('unselected' as const),
	});
	const all = ['a', 'b', 'c', 'd'].map((key) => item(key, true));
	expect(selectionAfterClick(all, 'b', 'single')).toEqual(['b']);
	const one = ['a', 'b', 'c', 'd'].map((key) => item(key, key === 'b'));
	expect(selectionAfterClick(one, 'b', 'single')).toBeNull();
	expect(selectionAfterClick(one, 'c', 'toggle')).toEqual(['b', 'c']);
	expect(selectionAfterClick(one, 'b', 'toggle')).toBeNull();
	expect(selectionAfterClick(one, 'd', 'range', 'b')).toEqual(['b', 'c', 'd']);
	expect(selectionAfterClick(one, 'd', 'single')).toEqual(['d']);
});
