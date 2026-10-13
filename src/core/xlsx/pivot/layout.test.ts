import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { getCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { styleAt } from '../styles';
import { pivotGrid } from './aggregate';

const cases = [
	['compact-top', 'compact', 'top', 8],
	['compact-bottom', 'compact', 'bottom', 10],
	['compact-bottom-no-subtotals', 'compact', 'bottom', 8],
	['outline-top', 'outline', 'top', 8],
	['outline-bottom', 'outline', 'bottom', 10],
	['outline-top-repeat', 'outline', 'top', 8],
] as const;

it.each([
	'custom indentation',
	'mixed field layout',
	'conflicting outline',
	'conflicting data outline',
	'malformed global layout',
])('preserves unsupported compact %s', async (kind) => {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('../__fixtures__/excel-pivot-compact-top.xlsx', import.meta.url)),
	);
	const path = 'xl/pivotTables/pivotTable1.xml';
	const original = await zip.file(path)!.async('string');
	const xml =
		kind === 'custom indentation'
			? original.replace('indent="0"', 'indent="2"')
			: kind === 'mixed field layout'
				? original.replace('<pivotField ', '<pivotField compact="0" ')
				: kind === 'conflicting outline'
					? original.replace('outline="1"', 'outline="0"')
					: kind === 'conflicting data outline'
						? original.replace('outlineData="1"', 'outlineData="0"')
						: original.replace('<pivotTableDefinition ', '<pivotTableDefinition compact="2" ');
	expect(xml).not.toBe(original);
	zip.file(path, xml);
	const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(wb.sheets.every((sheet) => !sheet.pivotTables)).toBe(true);
	expect(wb.warnings.some((warning) => warning.includes('cannot be refreshed'))).toBe(true);
	const saved = await JSZip.loadAsync(await saveXlsx(wb));
	expect(await saved.file(path)!.async('string')).toBe(xml);
});

for (const [name, layout, subtotalPosition, rows] of cases)
	it(`imports and refreshes native ${name} without losing its untouched parts`, async () => {
		const bytes = readFileSync(
			new URL(`../__fixtures__/excel-pivot-${name}.xlsx`, import.meta.url),
		);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const sheet = wb.sheets.findIndex((sheet) => sheet.name === 'Summary');
		const data = wb.sheets.findIndex((sheet) => sheet.name === 'Data');
		const pivot = wb.sheets[sheet]!.pivotTables![0]!;
		expect(pivot).toMatchObject({ layout, subtotalPosition });
		const grid = pivotGrid(pivot);
		expect(grid.values).toHaveLength(rows);
		expect(grid.values[0]).toEqual(
			layout === 'compact'
				? ['Row Labels', 'Sum of Amount']
				: ['Category', 'Region', 'Sum of Amount'],
		);
		expect(grid.values[1]?.at(-1)).toBe(subtotalPosition === 'top' ? 37 : null);
		const original = await JSZip.loadAsync(bytes);
		const untouched = await JSZip.loadAsync(await saveXlsx(wb));
		for (const path of Object.keys(original.files).filter((path) => path.startsWith('xl/pivot')))
			expect(await untouched.file(path)?.async('string')).toBe(
				await original.file(path)!.async('string'),
			);
		const edit = createEditSession(wb);
		edit.setCellInput(data, 1, 2, '100');
		edit.refreshPivotTable(sheet, pivot.name);
		const output = await loadXlsx(await saveXlsx(wb));
		expect(output.warnings).toEqual([]);
		expect(pivotGrid(output.sheets[sheet]!.pivotTables![0]!).values.at(-1)?.at(-1)).toBe(147);
		if (layout === 'compact') {
			expect(styleAt(output, getCell(output.sheets[sheet]!, 2, 0)?.styleId).alignment?.indent).toBe(
				1,
			);
			edit.refreshPivotTable(sheet, pivot.name, { layout: 'outline' });
			expect(styleAt(wb, getCell(wb.sheets[sheet]!, 2, 0)?.styleId).alignment?.indent ?? 0).toBe(0);
			edit.undo();
			expect(styleAt(wb, getCell(wb.sheets[sheet]!, 2, 0)?.styleId).alignment?.indent).toBe(1);
		}
	});

it('authors compact top headings, switches report widths atomically and preserves inactive top preference', async () => {
	const wb = await loadXlsx(
		readFileSync(new URL('../__fixtures__/excel-pivot-rows.xlsx', import.meta.url)),
	);
	const edit = createEditSession(wb);
	const sheet = edit.addSheet('Authored');
	const name = edit.createPivotTable(
		sheet,
		{ row: 0, col: 0 },
		{
			source: 'Data!A1:C6',
			rowField: 'Category',
			secondRowField: 'Region',
			valueField: 'Amount',
			layout: 'compact',
		},
	);
	const current = () => wb.sheets[sheet]!.pivotTables![0]!;
	expect(pivotGrid(current()).values).toEqual([
		['Row Labels', 'Sum of Amount'],
		['Fruit', 37],
		['East', 17],
		['West', 20],
		['Veg', 20],
		['East', 5],
		['West', 15],
		['Grand Total', 57],
	]);
	edit.refreshPivotTable(sheet, name, { layout: 'tabular' });
	expect(current().subtotalPosition).toBe('top');
	expect(pivotGrid(current()).values[3]).toEqual(['Fruit Total', null, 37]);
	expect(current().location.end.col).toBe(2);
	edit.undo();
	expect(current().layout).toBe('compact');
	edit.redo();
	expect(current().layout).toBe('tabular');
	const loaded = await loadXlsx(await saveXlsx(wb));
	expect(loaded.warnings).toEqual([]);
	expect(loaded.sheets[sheet]!.pivotTables![0]).toMatchObject({
		layout: 'tabular',
		subtotalPosition: 'top',
	});
	edit.refreshPivotTable(sheet, name, { layout: 'compact' });
	edit.setCellInput(sheet, 1, 2, 'note');
	expect(() => edit.refreshPivotTable(sheet, name, { layout: 'outline' })).toThrow('overwrite');
	expect(current().layout).toBe('compact');
});
