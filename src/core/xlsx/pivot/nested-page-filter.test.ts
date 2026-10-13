import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it, vi } from 'vitest';
import {
	createWorkbook,
	createEditSession,
	loadXlsx,
	saveXlsx,
	type PivotTableOptions,
} from '../index';
import { pivotGrid } from './aggregate';
import * as measures from './measures';
import { buildXml, children, elements, first, NS, parseXml } from '../../xml/index';
import { assertPageAxisItems } from './page-filter-admission';
import evidence from '../__fixtures__/native-pivot-nested-page-evidence.json';

const native = (file = evidence.files[1]!.file) =>
	readFileSync(new URL('../__fixtures__/' + file, import.meta.url));
const base: PivotTableOptions = {
	source: 'Data!A1:F8',
	rowField: 'Category',
	secondRowField: 'Region',
	valueField: 'Amount',
	additionalValues: [
		{ field: 'Quantity', aggregate: 'count', caption: 'Count of Quantity' },
		{ field: 'Amount', aggregate: 'average', caption: 'Average of Amount' },
	],
};
function workbook() {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] }),
		edit = createEditSession(wb);
	[evidence.source.fields, ...evidence.source.records].forEach((row, r) =>
		row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))),
	);
	return { wb, edit };
}
for (const sample of evidence.cases) {
	it(`matches owned native nested report filter ${sample.profile} through save/reopen`, async () => {
		const { wb, edit } = workbook();
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				...base,
				layout: sample.layout as NonNullable<PivotTableOptions['layout']>,
				subtotals: sample.subtotal !== 'none',
				subtotalPosition: sample.subtotal === 'top' ? 'top' : 'bottom',
				repeatLabels: 'repeatLabels' in sample ? sample.repeatLabels : false,
				...(sample.columns ? { columnField: 'Year' } : {}),
				pageFilter: {
					field: 'Market',
					...(sample.selection === 'All' ? {} : { item: sample.selection }),
				},
			},
		);
		const pivot = wb.sheets[1]!.pivotTables![0]!;
		expect(pivotGrid(pivot).values).toEqual(sample.values);
		expect(pivotGrid(pivot).indents).toEqual(sample.indents);
		expect(pivot.records).toHaveLength(7);
		const reopened = await loadXlsx(await saveXlsx(wb));
		expect(reopened.warnings).toEqual([]);
		expect(pivotGrid(reopened.sheets[1]!.pivotTables![0]!).values).toEqual(sample.values);
	});
}
for (const file of evidence.files) {
	it(`preserves and refreshes the complete native cache ${file.file}`, async () => {
		const bytes = native(file.file),
			wb = await loadXlsx(bytes),
			edit = createEditSession(wb);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		const expected = evidence.cases.find((p) => p.profile === file.profile)!;
		expect(pivotGrid(pivot).values).toEqual(expected.values);
		const original = await JSZip.loadAsync(bytes),
			saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(original.files).filter(
			(p) => p.startsWith('xl/pivot') && !original.files[p]!.dir,
		))
			expect(await saved.file(part)!.async('string')).toBe(
				await original.file(part)!.async('string'),
			);
		edit.refreshPivotTable(0, pivot.name, { pageFilter: { field: 'Market' } });
		expect(wb.sheets[0]!.pivotTables![0]!.records).toHaveLength(7);
		edit.undo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(expected.values);
		edit.redo();
		const reload = await loadXlsx(await saveXlsx(wb));
		expect(reload.warnings).toEqual([]);
		expect(reload.sheets[0]!.pivotTables![0]!.pageFilter).toEqual({ field: 'Market' });
	});
}
it('maps both selected hierarchy levels to the complete native item lists', async () => {
	const { wb, edit } = workbook();
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{ ...base, pageFilter: { field: 'Market', item: 'East' } },
	);
	const zip = await JSZip.loadAsync(await saveXlsx(wb));
	const root = parseXml(
		await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string'),
	).documentElement;
	const fields = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x);
	expect(children(first(fields[0], 'items', NS.x)!, 'item', NS.x)).toHaveLength(4);
	expect(children(first(fields[1], 'items', NS.x)!, 'item', NS.x)).toHaveLength(4);
	const row = children(first(root, 'rowItems', NS.x)!, 'i', NS.x)[0]!;
	expect(children(row, 'x', NS.x).map((x) => x.getAttribute('v') ?? '0')).toEqual(['1', '1']);
});
it.each(['Region', 'Category', 'Year', 'Amount'])(
	'refuses filter/axis collisions %s atomically',
	(field) => {
		const { wb, edit } = workbook();
		expect(() =>
			edit.createPivotTable(
				1,
				{ row: 0, col: 0 },
				{ ...base, columnField: 'Year', pageFilter: { field } },
			),
		).toThrow();
		expect(wb.sheets[1]!.pivotTables ?? []).toHaveLength(0);
	},
);
it('counts nested headings and two page rows before header allocation', () => {
	const { wb, edit } = workbook();
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{ ...base, layout: 'outline', pageFilter: { field: 'Market' } },
	);
	const pivot = wb.sheets[1]!.pivotTables![0]!;
	pivot.records = Array.from({ length: 27776 }, (_, i) => [i, 'North', 'East', 2025, 1, 1]);
	pivot.additionalValues = Array.from({ length: 15 }, (_, i) => ({
		field: 'Amount',
		aggregate: 'sum',
		caption: 'Measure ' + i,
	}));
	const headers = vi.spyOn(measures, 'pivotHeaders');
	try {
		expect(() => pivotGrid(pivot)).toThrow('too large');
		expect(headers).not.toHaveBeenCalled();
	} finally {
		headers.mockRestore();
	}
});
it.each(['prefix', 'inner-index', 'inner-count', 'row-count'])(
	'preserves unsupported nested report metadata %s',
	async (change) => {
		const zip = await JSZip.loadAsync(native()),
			part = 'xl/pivotTables/pivotTable1.xml';
		const root = parseXml(await zip.file(part)!.async('string')).documentElement;
		const rows = children(first(root, 'rowItems', NS.x)!, 'i', NS.x);
		if (change === 'prefix') rows[1]!.setAttribute('r', '2');
		if (change === 'inner-index') children(rows[0]!, 'x', NS.x)[1]!.setAttribute('v', '0');
		if (change === 'inner-count')
			first(
				children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x)[1],
				'items',
				NS.x,
			)!.setAttribute('count', '2');
		if (change === 'row-count') first(root, 'rowItems', NS.x)!.setAttribute('count', '999');
		const { buildXml } = await import('../../xml/index');
		const modified = buildXml(root);
		zip.file(part, modified);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await saved.file(part)!.async('string')).toBe(modified);
	},
);
it('preserves native repeated-column captions rather than discarding the extra header profile', async () => {
	const source = await JSZip.loadAsync(
		native('excel-pivot-nested-page-unsupported-column-repeat.xlsx'),
	);
	const wb = await loadXlsx(await source.generateAsync({ type: 'uint8array' }));
	expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
	expect(wb.warnings).toEqual([
		'PivotTable parts on "Summary" are preserved; this layout or data source cannot be refreshed or configured.',
	]);
	const saved = await JSZip.loadAsync(await saveXlsx(wb));
	for (const part of Object.keys(source.files).filter(
		(p) => p.startsWith('xl/pivot') && !source.files[p]!.dir,
	))
		expect(await saved.file(part)!.async('string')).toBe(await source.file(part)!.async('string'));
});
it.each(
	[
		'1',
		'true',
		' 1 ',
		' true ',
		'0',
		'false',
		' 0 ',
		' false ',
		'',
		'TRUE',
		'yes',
		'2',
		'0.0',
	].flatMap((value) => [false, true].map((later) => [value, later] as const)),
)('guards repeated-column value %j in a later extension %s', async (value, later) => {
	const bytes = native('excel-pivot-nested-page-tabular-east-bottom-columns-repeat.xlsx');
	const wb = await loadXlsx(bytes),
		pivot = wb.sheets[0]!.pivotTables![0]!;
	const source = await JSZip.loadAsync(bytes),
		part = 'xl/pivotTables/pivotTable1.xml';
	const root = parseXml(await source.file(part)!.async('string')).documentElement;
	const column = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x)[3]!;
	const list = parseXml(
		`<extLst xmlns="${NS.x}">${later ? '<ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}"/>' : ''}<ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}"><x14:pivotField xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" fillDownLabels="${value}"/></ext></extLst>`,
	).documentElement;
	column.appendChild(list);
	const cache = parseXml(
		await source.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string'),
	).documentElement;
	const records = parseXml(
		await source.file('xl/pivotCache/pivotCacheRecords1.xml')!.async('string'),
	).documentElement;
	const shared = children(first(cache, 'cacheFields', NS.x)!, 'cacheField', NS.x).map((field) =>
		elements(first(field, 'sharedItems', NS.x)!).map((item) =>
			item.localName === 'n' ? Number(item.getAttribute('v')) : item.getAttribute('v'),
		),
	);
	// Exercise this guard directly: generic extension admission is deliberately stricter (0/1).
	const inspect = () => assertPageAxisItems(root, pivot, shared, cache, records);
	if (['0', 'false'].includes(value.trim())) {
		expect(inspect).not.toThrow();
		if (value === '0') {
			const changed = buildXml(root);
			source.file(part, changed);
			const loaded = await loadXlsx(await source.generateAsync({ type: 'uint8array' }));
			expect(loaded.warnings).toEqual([]);
			expect(loaded.sheets[0]!.pivotTables).toHaveLength(1);
			const saved = await JSZip.loadAsync(await saveXlsx(loaded));
			expect(await saved.file(part)!.async('string')).toBe(changed);
		}
		return;
	}
	expect(inspect).toThrow(
		['1', 'true'].includes(value.trim())
			? 'Repeated report column labels'
			: 'Invalid report column label repetition',
	);
	const changed = buildXml(root);
	source.file(part, changed);
	const loaded = await loadXlsx(await source.generateAsync({ type: 'uint8array' }));
	expect(loaded.sheets[0]!.pivotTables ?? []).toHaveLength(0);
	const saved = await JSZip.loadAsync(await saveXlsx(loaded));
	for (const name of Object.keys(source.files).filter(
		(name) => name.startsWith('xl/pivot') && !source.files[name]!.dir,
	))
		expect(await saved.file(name)!.async('string')).toBe(await source.file(name)!.async('string'));
});
