import { readFileSync } from 'node:fs';
import { expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import type { PivotTable } from '../model';
import { pivotGrid } from './aggregate';
import * as columns from './columns';

for (const [suffix, layout] of [
	['compact', 'compact'],
	['outline', 'outline'],
	['compact-rows-top-columns', 'compact'],
	['outline-rows-bottom-columns', 'outline'],
] as const)
	it(`imports native multiple measures in ${suffix} and refreshes every measure`, async () => {
		const bytes = readFileSync(
			new URL(`../__fixtures__/excel-pivot-values-${suffix}.xlsx`, import.meta.url),
		);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.layout).toBe(layout);
		expect(pivot.additionalValues).toHaveLength(2);
		expect(pivotGrid(pivot).values.at(-1)?.slice(-3)).toEqual([57, 5, 11.4]);
		const original = await JSZip.loadAsync(bytes);
		const untouched = await JSZip.loadAsync(await saveXlsx(wb));
		for (const path of Object.keys(original.files).filter((path) => path.startsWith('xl/pivot')))
			expect(await untouched.file(path)?.async('string')).toBe(
				await original.file(path)!.async('string'),
			);
		const edit = createEditSession(wb);
		edit.setCellInput(1, 1, 2, '100');
		edit.refreshPivotTable(0, pivot.name);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(pivotGrid(loaded.sheets[0]!.pivotTables![0]!).values.at(-1)?.slice(-3)).toEqual([
			147, 5, 29.4,
		]);
		edit.undo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values.at(-1)?.slice(-3)).toEqual([
			57, 5, 11.4,
		]);
	});

it.each([false, true])(
	'counts all measure header rows in the million-cell limit (hierarchy %s)',
	(hierarchy) => {
		const pivot: PivotTable = {
			name: 'Large',
			source: 'A1:C10',
			fields: ['Category', 'Column', 'Amount', 'Inner'],
			rowField: 'Category',
			columnField: 'Column',
			valueField: 'Amount',
			valueCaption: 'Sum of Amount',
			aggregate: 'sum',
			layout: 'tabular',
			additionalValues: Array.from({ length: 15 }, (_, i) => ({
				field: 'Amount',
				aggregate: 'sum',
				caption: `Value ${i}`,
			})),
			location: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			records: Array.from({ length: hierarchy ? 8929 : 10417 }, (_, i) => [
				hierarchy ? 'Fruit' : `Group ${i % 2}`,
				i,
				1,
				`Inner ${i % 2}`,
			]),
			...(hierarchy ? { secondRowField: 'Inner' } : {}),
		};
		const headers = vi.spyOn(columns, 'pivotColumnHeaders');
		try {
			expect(() => pivotGrid(pivot)).toThrow('too large');
			expect(headers).not.toHaveBeenCalled();
		} finally {
			headers.mockRestore();
		}
	},
);

it('reads a custom Values caption and refuses a header that does not match it', async () => {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('../__fixtures__/excel-pivot-values-compact.xlsx', import.meta.url)),
	);
	const part = 'xl/pivotTables/pivotTable1.xml';
	const xml = (await zip.file(part)!.async('string')).replace(
		'dataCaption="Values"',
		'dataCaption="Metrics"',
	);
	zip.file(part, xml);
	const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(wb.warnings).toEqual([]);
	expect(wb.sheets[0]!.pivotTables![0]!.captions).toEqual({ values: 'Metrics' });
	expect(await (await JSZip.loadAsync(await saveXlsx(wb))).file(part)!.async('string')).toBe(xml);
	const single = await loadXlsx(
		readFileSync(new URL('../__fixtures__/excel-pivot.xlsx', import.meta.url)),
	);
	const pivot = single.sheets[0]!.pivotTables![0]!;
	pivot.sourceXml = pivot.sourceXml!.replace('dataCaption="Values"', 'dataCaption="Metrics"');
	expect(() =>
		createEditSession(single).refreshPivotTable(0, pivot.name, {
			additionalValues: [{ field: 'Amount', aggregate: 'count' }],
		}),
	).toThrow('unsupported Values header');
	expect(single.sheets[0]!.pivotTables![0]).toBe(pivot);
});
