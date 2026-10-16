import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it, vi } from 'vitest';
import { createWorkbook, createEditSession, loadXlsx, saveXlsx, getCell } from '../index';
import { pivotGrid } from './aggregate';
import * as columns from './columns';
import { children, first, parseXml, NS } from '../../xml/index';
const profiles = [
	'tabular-all',
	'compact-west',
	'outline-east',
	'tabular-number',
	'tabular-boolean',
	'tabular-blank',
	'tabular-error',
];
const native = (profile = 'tabular-all') =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-page-${profile}.xlsx`, import.meta.url));
for (const profile of profiles) {
	it(`imports native report filter ${profile} and preserves every source part`, async () => {
		const bytes = native(profile),
			wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.pageFilter?.field).toBe('Region');
		expect(pivot.location.start.row).toBe(0);
		expect(pivotGrid(pivot).values[0]?.[0]).toBe('Region');
		const source = await JSZip.loadAsync(bytes),
			saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(source.files).filter(
			(p) => p.startsWith('xl/pivot') && !source.files[p]!.dir,
		))
			expect(await saved.file(part)!.async('string')).toBe(
				await source.file(part)!.async('string'),
			);
	});
	it(`refreshes ${profile} with full cache, selection, history and save/reopen`, async () => {
		const wb = await loadXlsx(native(profile)),
			edit = createEditSession(wb);
		const pivot = wb.sheets[0]!.pivotTables![0]!,
			before = pivotGrid(pivot).values;
		edit.refreshPivotTable(0, pivot.name, { pageFilter: { field: 'Region' } });
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values[0]?.[1]).toBe('(All)');
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(loaded.sheets[0]!.pivotTables![0]!.records).toEqual(pivot.records);
		edit.undo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(before);
		edit.redo();
		expect(wb.sheets[0]!.pivotTables![0]!.pageFilter).toEqual({ field: 'Region' });
	});
}
it('resolves field-item indexes, retains hidden axes and clears report rows on removal', async () => {
	const wb = await loadXlsx(native('compact-west')),
		edit = createEditSession(wb),
		pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.pageFilter).toEqual({ field: 'Region', item: 'West' });
	expect(pivotGrid(pivot).columnItems).toEqual([2026]);
	edit.refreshPivotTable(0, pivot.name);
	const zip = await JSZip.loadAsync(await saveXlsx(wb));
	const xml = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string');
	expect(xml).toContain('axis="axisPage"');
	expect(xml).toContain('ref="A3:G8"');
	expect(xml).toContain('item="1"');
	expect(xml).toContain('<x v="1"/>');
	edit.refreshPivotTable(0, pivot.name, { pageFilter: null });
	expect(wb.sheets[0]!.pivotTables![0]!.pageFilter).toBeUndefined();
	expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values[0]?.[0]).not.toBe('Region');
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]!.pageFilter?.item).toBe('West');
});
it.each([
	null,
	[],
	{ field: 'Missing' },
	{ field: 'Category' },
	{ field: 'Amount' },
	{ field: 'Region', item: ['East'] },
	{ field: 'Region', item: { error: '#bad' } },
	{ field: 'Region', item: Infinity },
	{ field: 'Region', item: 'east' },
	{ field: 'Region', item: 'Missing' },
	{ field: 'Region', dynamic: true },
])('refuses unsupported runtime report filter %j atomically', async (filter) => {
	const wb = await loadXlsx(native()),
		edit = createEditSession(wb),
		pivot = wb.sheets[0]!.pivotTables![0]!;
	if (filter === null) return;
	expect(() => edit.refreshPivotTable(0, pivot.name, { pageFilter: filter as never })).toThrow();
	expect(edit.canUndo()).toBe(false);
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
});
it('keeps source and report-result cells protected from direct writes, and missing refresh items atomic', async () => {
	const wb = await loadXlsx(native('compact-west')),
		edit = createEditSession(wb),
		pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(() => edit.setCellInput(0, 0, 1, 'East')).toThrow('cannot be edited directly');
	for (const row of [2, 4]) edit.setCellInput(1, row, 1, 'East');
	const before = wb.sheets[0]!.pivotTables![0];
	expect(() => edit.refreshPivotTable(0, pivot.name)).toThrow('missing from the source');
	expect(wb.sheets[0]!.pivotTables![0]).toBe(before);
	expect(getCell(wb.sheets[0]!, 0, 1)?.value).toBe('West');
});
it.each(['item="2"', 'item="-1"', 'item="0.0"', 'item="NaN"', 'hier="0"', 'fld="2"'])(
	'preserves unsupported page selector metadata %s',
	async (change) => {
		const zip = await JSZip.loadAsync(native('compact-west')),
			part = 'xl/pivotTables/pivotTable1.xml';
		const old = await zip.file(part)!.async('string');
		const xml = old.replace(
			/<pageField fld="1" item="1" hier="-1"\/>/,
			`<pageField ${change.startsWith('fld=') ? '' : 'fld="1"'} ${change} ${change.startsWith('hier=') ? '' : 'hier="-1"'}/>`,
		);
		zip.file(part, xml);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await saved.file(part)!.async('string')).toBe(xml);
	},
);
it('includes the two report rows in the output budget before allocating headers', () => {
	const wb = createWorkbook(),
		edit = createEditSession(wb);
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 1],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	const name = edit.createPivotTable(
		0,
		{ row: 10, col: 0 },
		{
			source: 'A1:C2',
			rowField: 'Category',
			valueField: 'Amount',
			pageFilter: { field: 'Region' },
		},
	);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	pivot.records = Array.from({ length: 58821 }, (_, i) => [i, 'East', 1]);
	pivot.additionalValues = Array.from({ length: 15 }, (_, i) => ({
		field: 'Amount',
		aggregate: 'sum',
		caption: 'Measure ' + i,
	}));
	const headers = vi.spyOn(columns, 'pivotColumnHeaders');
	try {
		expect(() => pivotGrid(pivot)).toThrow('too large');
		expect(headers).not.toHaveBeenCalled();
	} finally {
		headers.mockRestore();
	}
	expect(name).toBe(pivot.name);
});

it('refuses Values on rows with a single measure behind a report filter before history', async () => {
	const wb = await loadXlsx(native()),
		edit = createEditSession(wb),
		pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(() => edit.refreshPivotTable(0, pivot.name, { valuesAxis: 'rows' })).toThrow(
		'Values on rows require multiple measures',
	);
	expect(edit.canUndo()).toBe(false);
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
});
it('refuses ambiguous typed item captions without changing source or report', () => {
	const wb = createWorkbook(),
		edit = createEditSession(wb);
	[
		['Category', 'Region', 'Amount'],
		['Fruit', true, 10],
		['Veg', 'TRUE', 20],
	].forEach((line, r) =>
		line.forEach((v, c) =>
			edit.setCellInput(0, r, c, typeof v === 'string' && v === 'TRUE' ? "'TRUE" : String(v)),
		),
	);
	const before = edit.canUndo();
	expect(() =>
		edit.createPivotTable(
			0,
			{ row: 10, col: 0 },
			{
				source: 'A1:C3',
				rowField: 'Category',
				valueField: 'Amount',
				pageFilter: { field: 'Region' },
			},
		),
	).toThrow('Ambiguous');
	expect(edit.canUndo()).toBe(before);
	expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
});
it.each(['count', 'location', 'multiple', 'cache-index', 'cache-count'])(
	'preserves unmodeled filter/cache metadata %s',
	async (mode) => {
		const zip = await JSZip.loadAsync(native('compact-west')),
			part = 'xl/pivotTables/pivotTable1.xml';
		let xml = await zip.file(part)!.async('string');
		if (mode === 'count') xml = xml.replace('<pageFields count="1">', '<pageFields count="2">');
		if (mode === 'location') xml = xml.replace('rowPageCount="1"', 'rowPageCount="2"');
		if (mode === 'multiple')
			xml = xml.replace('multipleFieldFilters="0"', 'multipleFieldFilters="1"');
		if (mode === 'cache-index')
			xml = xml.replace(/(<pivotField axis="axisPage"[^>]*><items count="3"><item x=")0/, '$199');
		if (mode === 'cache-count')
			xml = xml.replace(/(<pivotField axis="axisPage"[^>]*><items count=")3/, '$12');
		zip.file(part, xml);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await saved.file(part)!.async('string')).toBe(xml);
	},
);

it('maps visible column entries to complete cache items after a report selection', async () => {
	const wb = await loadXlsx(native('compact-west'));
	createEditSession(wb).refreshPivotTable(0, 'SalesPivot');
	const zip = await JSZip.loadAsync(await saveXlsx(wb)),
		root = parseXml(
			await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string'),
		).documentElement;
	const col = first(root, 'colItems', NS.x)!;
	expect(col.getAttribute('count')).toBe('6');
	expect(children(col, 'i', NS.x)).toHaveLength(6);
	expect(children(children(col, 'i', NS.x)[0]!, 'x', NS.x)[0]!.getAttribute('v')).toBe('1');
	const loaded = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(loaded.warnings).toEqual([]);
});
it('counts visible row entries independently of complete source field items for sparse reports', async () => {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] }),
		edit = createEditSession(wb);
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Veg', 'North', 5],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{
			source: 'Data!A1:C3',
			rowField: 'Category',
			valueField: 'Amount',
			pageFilter: { field: 'Region', item: 'East' },
		},
	);
	const zip = await JSZip.loadAsync(await saveXlsx(wb)),
		root = parseXml(
			await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string'),
		).documentElement;
	expect(first(root, 'rowItems', NS.x)!.getAttribute('count')).toBe('2');
	expect(children(first(root, 'rowItems', NS.x)!, 'i', NS.x)).toHaveLength(2);
	expect(
		first(
			children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x)[0],
			'items',
			NS.x,
		)!.getAttribute('count'),
	).toBe('3');
	const loaded = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(loaded.warnings).toEqual([]);
	expect(pivotGrid(loaded.sheets[1]!.pivotTables![0]!).rowItems).toEqual(['Fruit']);
});

it.each(['boolean-alias', 'record-index'])(
	'preserves unmodeled cache lexical metadata %s',
	async (mode) => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] }),
			edit = createEditSession(wb);
		[
			['Category', 'Region', 'Amount'],
			['Fruit', true, 10],
			['Veg', false, 20],
		].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				source: 'Data!A1:C3',
				rowField: 'Category',
				valueField: 'Amount',
				pageFilter: { field: 'Region' },
			},
		);
		const zip = await JSZip.loadAsync(await saveXlsx(wb)),
			part =
				mode === 'boolean-alias'
					? 'xl/pivotCache/pivotCacheDefinition1.xml'
					: 'xl/pivotCache/pivotCacheRecords1.xml';
		const xml = (await zip.file(part)!.async('string')).replace(
			mode === 'boolean-alias' ? '<b v="1"/>' : '<x v="0"/>',
			mode === 'boolean-alias' ? '<b v="true"/>' : '<x v="0.0"/>',
		);
		zip.file(part, xml);
		const loaded = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(loaded.sheets[1]!.pivotTables ?? []).toHaveLength(0);
		const saved = await JSZip.loadAsync(await saveXlsx(loaded));
		expect(await saved.file(part)!.async('string')).toBe(xml);
	},
);
