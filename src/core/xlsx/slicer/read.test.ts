import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-slicer-${name}.xlsx`, import.meta.url));

it('reads table slicers, their caches and the table column filter', async () => {
	const wb = await loadXlsx(fixture('table'));
	expect(wb.warnings).toEqual([]);
	expect(wb.slicers?.slicers.map((s) => [s.name, s.cache, s.columnCount, s.rowHeight])).toEqual([
		['Category', 'Slicer_Category', undefined, 251883],
		['Region', 'Slicer_Region', 2, 251883],
	]);
	const caches = [...wb.slicers!.caches].sort((a, b) => a.name.localeCompare(b.name));
	expect(caches.map((c) => [c.name, c.sourceName, c.source])).toEqual([
		['Slicer_Category', 'Category', { kind: 'table', tableId: 1, column: 1 }],
		['Slicer_Region', 'Region', { kind: 'table', tableId: 1, column: 2 }],
	]);
	const sheet = wb.sheets[0]!;
	expect(sheet.drawings.map((d) => d.kind)).toEqual(['slicer', 'slicer']);
	expect(sheet.tables[0]!.columns[0]!.filter).toEqual({ values: ['Fruit', 'Nuts'] });
	expect(sheet.rowInfo.get(3)).toMatchObject({ hidden: true, filteredOut: true });
});

it('reads a slicer connected to a PivotTable and the hidden item', async () => {
	const wb = await loadXlsx(fixture('pivot'));
	expect(wb.warnings).toEqual([]);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.itemFilters).toEqual([{ field: 'Category', hidden: ['Veg'] }]);
	expect(pivot.slicerCacheId).toBeTypeOf('number');
	expect(wb.slicers?.caches[0]?.source).toMatchObject({ kind: 'pivot' });
});

for (const kind of ['table', 'pivot', 'timeline']) {
	it(`round-trips unchanged ${kind} slicer parts byte for byte`, async () => {
		const bytes = fixture(kind);
		const original = await JSZip.loadAsync(bytes);
		const saved = await JSZip.loadAsync(await saveXlsx(await loadXlsx(bytes)));
		const names = Object.keys(original.files).filter((name) =>
			/slicer|timeline|drawings\/drawing/.test(name),
		);
		expect(names.length).toBeGreaterThan(2);
		for (const name of names)
			expect(await saved.file(name)?.async('string'), name).toBe(
				await original.file(name)!.async('string'),
			);
		const sheet = await saved.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(sheet).toMatch(/slicerList|timelineRefs/);
		const book = await saved.file('xl/workbook.xml')!.async('string');
		expect(book).toMatch(/slicerCache|timelineCacheRef/);
	});
}

it('keeps a timeline as a placeholder until its date field is supported', async () => {
	const wb = await loadXlsx(fixture('timeline'));
	expect(wb.sheets[0]!.drawings[0]).toMatchObject({ kind: 'unsupported', description: 'timeline' });
});
