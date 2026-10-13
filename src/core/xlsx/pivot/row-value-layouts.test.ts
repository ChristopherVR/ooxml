import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it, vi } from 'vitest';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { pivotGrid } from './aggregate';
import * as measures from './measures';

for (const profile of [
	'compact',
	'compact-columns',
	'outline',
	'outline-columns',
	'outline-columns-repeat',
])
	it(`imports native row measures in ${profile}, preserves, refreshes and reopens`, async () => {
		const bytes = readFileSync(
			new URL(`../__fixtures__/excel-pivot-row-values-${profile}.xlsx`, import.meta.url),
		);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		const grid = pivotGrid(pivot);
		expect(pivot.valuesAxis).toBe('rows');
		expect(pivot.layout).toBe(profile.split('-')[0]);
		expect(grid.values.length).toBe(profile.includes('columns') ? 13 : 12);
		expect(grid.rowLabelColumns).toBe(profile.startsWith('compact') ? 1 : 2);
		expect(grid.values.slice(-3).map((row) => row.at(-1))).toEqual([57, 5, 11.4]);
		const source = await JSZip.loadAsync(bytes);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(source.files).filter((part) => part.startsWith('xl/pivot')))
			expect(await saved.file(part)?.async('string')).toBe(
				await source.file(part)!.async('string'),
			);
		if (profile.startsWith('compact'))
			expect(grid.indents?.filter((indent) => indent === 1)).toHaveLength(6);
		if (profile.endsWith('repeat')) expect(grid.values[3]?.[0]).toBe('Fruit');
		const edit = createEditSession(wb);
		edit.setCellInput(1, 1, 2, '100');
		edit.refreshPivotTable(0, pivot.name);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(
			pivotGrid(loaded.sheets[0]!.pivotTables![0]!)
				.values.slice(-3)
				.map((row) => row.at(-1)),
		).toEqual([147, 5, 29.4]);
		edit.undo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(grid.values);
	});

it('converts every Values-row layout with undo and removes compact indent styles on conversion', async () => {
	const wb = await loadXlsx(
		readFileSync(
			new URL('../__fixtures__/excel-pivot-row-values-compact-columns.xlsx', import.meta.url),
		),
	);
	const edit = createEditSession(wb);
	const current = () => wb.sheets[0]!.pivotTables![0]!;
	for (const layout of ['outline', 'tabular', 'compact'] as const) {
		edit.refreshPivotTable(0, current().name, { layout });
		const saved = await loadXlsx(await saveXlsx(wb));
		expect(saved.warnings).toEqual([]);
		expect(saved.sheets[0]!.pivotTables![0]?.layout).toBe(layout);
		expect(
			pivotGrid(current())
				.values.slice(-3)
				.map((row) => row.at(-1)),
		).toEqual([57, 5, 11.4]);
	}
	edit.undo();
	expect(current().layout).toBe('tabular');
	edit.undo();
	expect(current().layout).toBe('outline');
	edit.undo();
	expect(current().layout).toBe('compact');
});

it('counts blank group headings before allocating compact Values-row output', async () => {
	const wb = await loadXlsx(
		readFileSync(
			new URL('../__fixtures__/excel-pivot-row-values-compact-columns.xlsx', import.meta.url),
		),
	);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	pivot.additionalValues = Array.from({ length: 15 }, (_, i) => ({
		field: 'Amount',
		aggregate: 'sum',
		caption: `Measure ${i}`,
	}));
	pivot.records = Array.from({ length: 19229 }, (_, i) => [
		i % 2 ? 'Fruit' : 'Veg',
		'East',
		1,
		1,
		i,
	]);
	// 52 rows (including two headings, two headers and 16 totals) * 19,231 columns.
	const headers = vi.spyOn(measures, 'pivotHeaders');
	try {
		expect(() => pivotGrid(pivot)).toThrow('too large');
		expect(headers).not.toHaveBeenCalled();
	} finally {
		headers.mockRestore();
	}
});

it('preserves custom compact indentation without accepting an editable profile', async () => {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('../__fixtures__/excel-pivot-row-values-compact.xlsx', import.meta.url)),
	);
	const part = 'xl/pivotTables/pivotTable1.xml';
	const original = await zip.file(part)!.async('string');
	const xml = original.replace('indent="0"', 'indent="2"');
	expect(xml).not.toBe(original);
	zip.file(part, xml);
	const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(wb.sheets[0]!.pivotTables).toBeUndefined();
	expect(wb.warnings.some((w) => w.includes('cannot be refreshed'))).toBe(true);
	expect(await (await JSZip.loadAsync(await saveXlsx(wb))).file(part)!.async('string')).toBe(xml);
});
