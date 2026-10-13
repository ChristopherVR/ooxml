import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { pivotGrid } from './aggregate';

for (const suffix of ['', '-columns', '-rows-columns'])
	it(`imports native multiple values ${suffix}, preserves parts and refreshes all measures`, async () => {
		const bytes = readFileSync(
			new URL(`../__fixtures__/excel-pivot-values${suffix}.xlsx`, import.meta.url),
		);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const summary = wb.sheets.findIndex((sheet) => sheet.name === 'Summary');
		const data = wb.sheets.findIndex((sheet) => sheet.name === 'Data');
		const pivot = wb.sheets[summary]!.pivotTables![0]!;
		expect(pivot.additionalValues).toEqual([
			{ field: 'Quantity', aggregate: 'count', caption: 'Count of Quantity' },
			{ field: 'Amount', aggregate: 'average', caption: 'Average of Amount' },
		]);
		expect(pivotGrid(pivot).values.at(-1)?.slice(-3)).toEqual([57, 5, 11.4]);
		const source = await JSZip.loadAsync(bytes);
		const untouched = await JSZip.loadAsync(await saveXlsx(wb));
		for (const path of Object.keys(source.files).filter((path) => path.startsWith('xl/pivot')))
			expect(await untouched.file(path)?.async('string')).toBe(
				await source.file(path)!.async('string'),
			);
		const edit = createEditSession(wb);
		edit.setCellInput(data, 1, 2, '100');
		edit.refreshPivotTable(summary, pivot.name);
		expect(pivotGrid(wb.sheets[summary]!.pivotTables![0]!).values.at(-1)?.slice(-3)).toEqual([
			147, 5, 29.4,
		]);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(loaded.sheets[summary]!.pivotTables![0]?.additionalValues).toEqual(
			pivot.additionalValues,
		);
		edit.undo();
		expect(pivotGrid(wb.sheets[summary]!.pivotTables![0]!).values.at(-1)?.slice(-3)).toEqual([
			57, 5, 11.4,
		]);
	});

it('switches measure counts atomically and refuses unsupported layouts or duplicate captions', async () => {
	const wb = await loadXlsx(
		readFileSync(new URL('../__fixtures__/excel-pivot-values.xlsx', import.meta.url)),
	);
	const edit = createEditSession(wb);
	const summary = wb.sheets.findIndex((sheet) => sheet.name === 'Summary');
	const name = wb.sheets[summary]!.pivotTables![0]!.name;
	expect(() => edit.refreshPivotTable(summary, name, { layout: 'compact' })).toThrow('tabular');
	expect(() =>
		edit.refreshPivotTable(summary, name, {
			additionalValues: [{ field: 'Amount', caption: 'Sum of Amount' }],
		}),
	).toThrow('unique');
	expect(() =>
		edit.refreshPivotTable(summary, name, {
			additionalValues: [{ field: 'Quantity', caption: 'Quantity' }],
		}),
	).toThrow('unique');
	expect(() =>
		edit.refreshPivotTable(summary, name, {
			additionalValues: [{ field: 'Quantity', aggregate: 'invalid' as never }],
		}),
	).toThrow('aggregate');
	expect(() => edit.refreshPivotTable(summary, name, { additionalValues: [{} as never] })).toThrow(
		'valid',
	);
	edit.refreshPivotTable(summary, name, { additionalValues: [] });
	expect(pivotGrid(wb.sheets[summary]!.pivotTables![0]!).values[0]).toEqual([
		'Category',
		'Sum of Amount',
	]);
	edit.undo();
	expect(wb.sheets[summary]!.pivotTables![0]?.additionalValues).toHaveLength(2);
	edit.redo();
	expect(wb.sheets[summary]!.pivotTables![0]?.additionalValues).toBeUndefined();
});

it.each(['visible', 'omitted'])(
	'sets the native hideValuesRow cache while converting single %s flag to multiple',
	async (mode) => {
		const wb = await loadXlsx(
			readFileSync(new URL('../__fixtures__/excel-pivot-visible-values.xlsx', import.meta.url)),
		);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		if (mode === 'omitted') pivot.sourceXml = pivot.sourceXml!.replace('hideValuesRow="0"', '');
		createEditSession(wb).refreshPivotTable(0, pivot.name, {
			additionalValues: [{ field: 'Amount', aggregate: 'count' }],
		});
		const saved = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file(pivot.partName!)!.async('string')).toContain('hideValuesRow="1"');
		expect((await loadXlsx(saved)).warnings).toEqual([]);
	},
);

it.each([
	'measure index',
	'Values on rows',
	'custom measure calculation',
	'visible Values row',
	'omitted Values row flag',
])('preserves unsupported %s without exposing refresh', async (kind) => {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('../__fixtures__/excel-pivot-values.xlsx', import.meta.url)),
	);
	const part = 'xl/pivotTables/pivotTable1.xml';
	const source = await zip.file(part)!.async('string');
	const xml =
		kind === 'measure index'
			? source.replace('i="1"', 'i="99"')
			: kind === 'Values on rows'
				? source.replace('<field x="0"/>', '<field x="-2"/>')
				: kind === 'visible Values row'
					? source.replace('hideValuesRow="1"', 'hideValuesRow="0"')
					: kind === 'omitted Values row flag'
						? source.replace('hideValuesRow="1"', '')
						: source.replace('<dataField ', '<dataField showDataAs="percentOfTotal" ');
	expect(xml).not.toBe(source);
	zip.file(part, xml);
	const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(wb.sheets.every((sheet) => !sheet.pivotTables)).toBe(true);
	expect(wb.warnings.some((warning) => warning.includes('cannot be refreshed'))).toBe(true);
	expect(await (await JSZip.loadAsync(await saveXlsx(wb))).file(part)!.async('string')).toBe(xml);
});
