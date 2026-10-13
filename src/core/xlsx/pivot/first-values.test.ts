import { readFileSync } from 'node:fs';
import { expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { pivotGrid } from './aggregate';
import * as measures from './measures';
import { elements, first, parseXml, NS } from '../../xml/index';

const file = (profile: string) =>
	profile === 'tabular'
		? 'excel-pivot-row-values-rows-first-columns.xlsx'
		: `excel-pivot-row-values-${profile.startsWith('compact') ? 'compact' : 'outline'}-rows-bottom-first-columns${profile.endsWith('repeat') ? '-repeat' : ''}.xlsx`;
const native = (profile = 'tabular') =>
	readFileSync(new URL(`../__fixtures__/${file(profile)}`, import.meta.url));

for (const profile of ['tabular', 'compact', 'outline', 'outline-repeat']) {
	it(`imports ${profile} first-field Values and preserves source XML`, async () => {
		const bytes = native(profile),
			wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.valuesPosition).toBe('first');
		const grid = pivotGrid(pivot);
		expect(grid.values).toHaveLength(profile === 'tabular' ? 23 : 32);
		expect(grid.values.at(-1)?.at(-1)).toBe(11.4);
		if (profile !== 'compact')
			expect(grid.values[1]?.slice(0, 3)).toEqual(['Values', 'Category', 'Region']);
		const source = await JSZip.loadAsync(bytes),
			saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(source.files).filter(
			(name) => name.startsWith('xl/pivot') && !source.files[name]!.dir,
		))
			expect(await saved.file(part)?.async('string')).toBe(
				await source.file(part)!.async('string'),
			);
	});
	it(`refreshes ${profile} source measures with save/reopen and history`, async () => {
		const wb = await loadXlsx(native(profile)),
			edit = createEditSession(wb);
		const original = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		edit.setCellInput(1, 1, 2, '100');
		edit.refreshPivotTable(0, 'SalesPivot');
		const expected = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(expected.at(-1)?.at(-1)).toBe(29.4);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(pivotGrid(loaded.sheets[0]!.pivotTables![0]!).values).toEqual(expected);
		edit.undo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(original);
		edit.redo();
		expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(expected);
	});
}

it.each([
	{ patch: { valuesPosition: 'last' as const } },
	{ patch: { valuesAxis: 'columns' as const } },
	{ patch: { secondRowField: '' } },
	{ patch: { valuesAxis: 'columns' as const, additionalValues: [] } },
])('normalizes first metadata on conversion $patch', async ({ patch }) => {
	const wb = await loadXlsx(native()),
		edit = createEditSession(wb);
	edit.refreshPivotTable(0, 'SalesPivot', patch);
	const bytes = await saveXlsx(wb),
		zip = await JSZip.loadAsync(bytes);
	expect(await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string')).not.toContain(
		'dataPosition=',
	);
	const loaded = await loadXlsx(bytes);
	expect(loaded.warnings).toEqual([]);
	expect(loaded.sheets[0]!.pivotTables![0]?.valuesPosition).not.toBe('first');
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]?.valuesPosition).toBe('first');
});

it.each(['unsupported', '', null, 1, false])(
	'refuses malformed public Values position %s atomically',
	async (position) => {
		const wb = await loadXlsx(native()),
			edit = createEditSession(wb),
			pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(() =>
			edit.refreshPivotTable(0, pivot.name, { valuesPosition: position as 'last' }),
		).toThrow('valid Values position');
		expect(edit.canUndo()).toBe(false);
		expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
	},
);

it.each([
	{ valuesAxis: 'columns' as const, valuesPosition: 'first' as const },
	{ secondRowField: '', valuesPosition: 'first' as const },
])(
	'refuses incompatible explicit first options $valuesAxis $secondRowField without history',
	async (patch) => {
		const wb = await loadXlsx(native()),
			edit = createEditSession(wb),
			pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(() => edit.refreshPivotTable(0, pivot.name, patch)).toThrow('require two row fields');
		expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
		expect(edit.canUndo()).toBe(false);
	},
);

it.each(['tabular', 'compact', 'outline'] as const)(
	'guards %s first headings before output allocation',
	async (layout) => {
		const wb = await loadXlsx(native()),
			pivot = wb.sheets[0]!.pivotTables![0]!;
		pivot.layout = layout;
		delete pivot.columnField;
		pivot.additionalValues = Array.from({ length: 15 }, (_, i) => ({
			field: 'Amount',
			aggregate: 'sum' as const,
			caption: `Measure ${i}`,
		}));
		pivot.records = Array.from({ length: 35_000 }, (_, i) => [i, 'East', 1, 2, 2025]);
		const headers = vi.spyOn(measures, 'pivotHeaders');
		try {
			expect(() => pivotGrid(pivot)).toThrow('too large');
			expect(headers).not.toHaveBeenCalled();
		} finally {
			headers.mockRestore();
		}
	},
);

it.each([undefined, '', '1', '2', '0.5', '0.0', 'NaN'])(
	'preserves inconsistent native first dataPosition %s',
	async (position) => {
		const source = await JSZip.loadAsync(native()),
			part = 'xl/pivotTables/pivotTable1.xml';
		const xml = (await source.file(part)!.async('string')).replace(
			'dataPosition="0"',
			position === undefined ? '' : `dataPosition="${position}"`,
		);
		source.file(part, xml);
		const wb = await loadXlsx(await source.generateAsync({ type: 'uint8array' }));
		expect(wb.sheets[0]!.pivotTables ?? []).toHaveLength(0);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await saved.file(part)!.async('string')).toBe(xml);
	},
);

it('serializes native first Values subtotal prefix and retained measure-major order', async () => {
	const wb = await loadXlsx(native());
	createEditSession(wb).refreshPivotTable(0, 'SalesPivot');
	const zip = await JSZip.loadAsync(await saveXlsx(wb));
	const root = parseXml(
		await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string'),
	).documentElement;
	expect(elements(first(root, 'rowFields', NS.x)!).map((n) => n.getAttribute('x'))).toEqual([
		'-2',
		'0',
		'1',
	]);
	const entries = elements(first(root, 'rowItems', NS.x)!);
	expect(entries[2]?.getAttribute('t')).toBe('default');
	expect(entries[2]?.getAttribute('r')).toBe('1');
	expect(elements(entries[2]!).map((n) => n.getAttribute('v') ?? '0')).toEqual(['0']);
	expect(entries[3]?.getAttribute('r')).toBe('1');
	expect(elements(entries[3]!).map((n) => n.getAttribute('v') ?? '0')).toEqual(['1', '0']);
	expect(entries[6]?.getAttribute('i')).toBe('1');
	expect(elements(entries[6]!).map((n) => n.getAttribute('v') ?? '0')).toEqual(['1', '0', '0']);
});
