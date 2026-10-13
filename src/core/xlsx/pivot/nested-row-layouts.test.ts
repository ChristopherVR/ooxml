import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { getCell } from '../cells';
import { styleAt } from '../styles';
import { pivotGrid } from './aggregate';

for (const profile of ['compact', 'outline', 'outline-repeat']) {
	const layout = profile.startsWith('compact') ? 'compact' : 'outline';
	const repeated = profile.endsWith('repeat');
	const native = () =>
		readFileSync(
			new URL(
				`../__fixtures__/excel-pivot-row-values-${layout}-rows-bottom-columns${repeated ? '-repeat' : ''}.xlsx`,
				import.meta.url,
			),
		);
	it(`imports native nested ${profile} row Values and retains its source parts`, async () => {
		const bytes = native(),
			wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot).toMatchObject({
			layout,
			valuesAxis: 'rows',
			secondRowField: 'Region',
			repeatLabels: repeated,
		});
		const grid = pivotGrid(pivot);
		expect(grid.values).toHaveLength(29);
		expect(grid.values[4]).toEqual(
			layout === 'compact'
				? ['Sum of Amount', 10, 7, 17]
				: [repeated ? 'Fruit' : null, repeated ? 'East' : null, 'Sum of Amount', 10, 7, 17],
		);
		expect(grid.indents?.slice(0, 7)).toEqual(
			layout === 'compact' ? [0, 0, 0, 1, 2, 2, 2] : [0, 0, 0, 0, 0, 0, 0],
		);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(before.files).filter(
			(p) => p.startsWith('xl/pivot') && !before.files[p]!.dir,
		))
			expect(await after.file(part)?.async('string')).toBe(
				await before.file(part)!.async('string'),
			);
	});
	it(`refreshes ${profile} nested rows, retaining native indentation through save and history`, async () => {
		const wb = await loadXlsx(native()),
			edit = createEditSession(wb);
		const initial = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		edit.setCellInput(1, 1, 2, '100');
		edit.refreshPivotTable(0, 'SalesPivot');
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!);
		expect(grid.values.at(-1)?.at(-1)).toBe(29.4);
		if (layout === 'compact') {
			expect(styleAt(wb, getCell(wb.sheets[0]!, 4, 0)?.styleId).alignment).toMatchObject({
				indent: 2,
				horizontal: 'left',
			});
			expect(styleAt(wb, getCell(wb.sheets[0]!, 3, 0)?.styleId).alignment?.indent).toBe(1);
		}
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(pivotGrid(loaded.sheets[0]!.pivotTables![0]!).values).toEqual(grid.values);
		edit.undo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(initial);
		edit.redo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(grid.values);
	});
}

it.each([{ order: [-2, 0, 1] }, { order: [0, -2, 1] }])(
	'preserves unmodeled Values placement $order',
	async ({ order }) => {
		const source = await JSZip.loadAsync(
			readFileSync(
				new URL(
					'../__fixtures__/excel-pivot-row-values-compact-rows-bottom-columns.xlsx',
					import.meta.url,
				),
			),
		);
		const part = 'xl/pivotTables/pivotTable1.xml';
		const xml = (await source.file(part)!.async('string')).replace(
			/<rowFields[^>]*>.*?<\/rowFields>/,
			`<rowFields count="3">${order.map((index) => `<field x="${index}"/>`).join('')}</rowFields>`,
		);
		source.file(part, xml);
		const wb = await loadXlsx(await source.generateAsync({ type: 'uint8array' }));
		expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
		expect(wb.warnings.some((warning) => warning.includes('preserved'))).toBe(true);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const name of Object.keys(source.files).filter(
			(name) => name.startsWith('xl/pivot') && !source.files[name]!.dir,
		))
			expect(await saved.file(name)?.async('string')).toBe(
				await source.file(name)!.async('string'),
			);
	},
);
