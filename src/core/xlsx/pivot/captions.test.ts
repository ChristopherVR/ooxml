import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildXml, children, first, NS, parseXml } from '../../xml/index';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';
import { ENGLISH_PIVOT_CAPTIONS, localizedPivotCaptions, pivotCaptions } from './captions';
import { PIVOT_SUBTOTALS } from './subtotals';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-${name}.xlsx`, import.meta.url));
const definition = async (bytes: Uint8Array) =>
	parseXml(
		await (await JSZip.loadAsync(bytes)).file('xl/pivotTables/pivotTable1.xml')!.async('string'),
	).documentElement;

/**
 * The fixtures are Excel files with captions in other languages. Only the English Excel is at
 * hand: it wrote the header, grand total, Values and data field names the way a localized Excel
 * stores them (`generate-excel-pivot-captions.ps1`), and the words only a UI language supplies
 * (subtotal rows, measure totals, `(All)`) were relabelled in the shared strings.
 */
const CASES = [
	{
		file: 'captions-de-compact',
		captions: {
			rowLabels: 'Zeilenbeschriftungen',
			grandTotal: 'Gesamtergebnis',
			subtotal: '{0} Ergebnis',
			aggregates: { sum: 'Summe von {0}' },
		},
		cells: ['Zeilenbeschriftungen', 'Summe von Amount', 'Fruit Ergebnis', 'Gesamtergebnis'],
	},
	{
		file: 'captions-de-columns',
		captions: {
			rowLabels: 'Zeilenbeschriftungen',
			columnLabels: 'Spaltenbeschriftungen',
			grandTotal: 'Gesamtergebnis',
			values: 'Werte',
			measureTotal: 'Gesamt {0}',
			aggregates: { sum: 'Summe von {0}', count: 'Anzahl von {0}' },
		},
		cells: ['Spaltenbeschriftungen', 'Gesamt Summe von Amount', 'Gesamt Anzahl von Amount'],
	},
	{
		file: 'captions-es-tabular-page',
		captions: {
			grandTotal: 'Total general',
			all: '(Todas)',
			aggregates: { sum: 'Suma de {0}' },
		},
		cells: ['(Todas)', 'Total general', 'Fruit Total'],
	},
	{
		file: 'captions-fr-rows-values',
		captions: {
			grandTotal: 'Total général',
			values: 'Valeurs',
			aggregates: { sum: 'Somme de {0}', count: 'Nombre de {0}' },
		},
		cells: ['Valeurs', 'Total Somme de Amount', 'Fruit Nombre de Amount'],
	},
	{
		file: 'captions-fr-rows-values-compact-page',
		captions: {
			rowLabels: 'Étiquettes de lignes',
			grandTotal: 'Total général',
			values: 'Valeurs',
			all: '(Tous)',
			aggregates: { sum: 'Somme de {0}', count: 'Nombre de {0}' },
		},
		cells: ['(Tous)', 'Étiquettes de lignes', 'Total Nombre de Amount'],
	},
	{
		file: 'captions-zh-compact',
		captions: {
			rowLabels: '行标签',
			columnLabels: '列标签',
			grandTotal: '总计',
			subtotal: '{0} 汇总',
			aggregates: { sum: '求和项:{0}' },
		},
		cells: ['列标签', '行标签', 'Fruit 汇总', '总计'],
	},
] as const;

describe('PivotTable captions in the language of the file', () => {
	for (const sample of CASES) {
		it(`reads, keeps and refreshes the captions of ${sample.file}`, async () => {
			const bytes = fixture(sample.file);
			const wb = await loadXlsx(bytes);
			expect(wb.warnings).toEqual([]);
			const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
			const pivot = wb.sheets[s]!.pivotTables![0]!;
			expect(pivot.captions).toMatchObject(sample.captions);
			const grid = pivotGrid(pivot);
			const flat = grid.values.flat();
			for (const text of sample.cells) expect(flat).toContain(text);
			expect(flat).not.toContain('Grand Total');
			expect(flat).not.toContain('Row Labels');

			// An untouched load keeps every pivot part byte for byte.
			const original = await JSZip.loadAsync(bytes);
			const saved = await JSZip.loadAsync(await saveXlsx(wb));
			for (const part of Object.keys(original.files).filter((p) => p.startsWith('xl/pivot')))
				expect(await saved.file(part)?.async('string')).toBe(
					await original.file(part)!.async('string'),
				);

			// A refresh after a source change writes the same words, not English.
			const edit = createEditSession(wb);
			const data = wb.sheets.findIndex((sheet) => sheet.name === 'Data');
			edit.setCellInput(data, 1, 4, '100');
			edit.refreshPivotTable(s, pivot.name);
			const refreshed = wb.sheets[s]!.pivotTables![0]!;
			const after = pivotGrid(refreshed).values;
			for (const text of sample.cells) expect(after.flat()).toContain(text);
			const reopened = await loadXlsx(await saveXlsx(wb));
			expect(reopened.warnings).toEqual([]);
			const again = reopened.sheets[s]!.pivotTables![0]!;
			expect(pivotGrid(again).values).toEqual(after);
			expect(again.captions).toEqual(refreshed.captions);
			// The cells on the sheet are the grid.
			after.forEach((line, r) =>
				line.forEach((value, c) =>
					expect(
						getCell(reopened.sheets[s]!, again.location.start.row + r, again.location.start.col + c)
							?.value ?? null,
					).toEqual(value),
				),
			);
			edit.undo();
			edit.undo();
		});
	}

	it.each(['subtotals-all', 'colsubs-rows-custom'])(
		'reads the custom subtotal words of a German %s workbook',
		async (name) => {
			// Relabelled the way a German Excel words subtotal rows and columns.
			const zip = await JSZip.loadAsync(fixture(name));
			const words: Record<string, string> = {
				Sum: 'Summe',
				Count: 'Anzahl',
				Average: 'Mittelwert',
				StdDevp: 'Standardabweichung (Grundgesamtheit)',
				StdDev: 'Standardabweichung',
				Varp: 'Varianz (Grundgesamtheit)',
				Var: 'Varianz',
				Product: 'Produkt',
			};
			let text = await zip.file('xl/sharedStrings.xml')!.async('string');
			text = text.replace(
				/(<t[^>]*>)(Fruit|Veg|2025|2026) (\w+)(<\/t>)/g,
				(_all, open: string, item: string, word: string, close: string) =>
					`${open}${item} ${words[word] ?? word}${close}`,
			);
			text = text.replace('>Grand Total<', '>Gesamtergebnis<');
			zip.file('xl/sharedStrings.xml', text);
			const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
			expect(wb.warnings).toEqual([]);
			const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
			const pivot = wb.sheets[s]!.pivotTables![0]!;
			expect(pivot.captions!.grandTotal).toBe('Gesamtergebnis');
			expect(pivot.captions!.functions!.sum).toBe('Summe');
			expect(pivotGrid(pivot).values.flat()).toContain(
				name === 'subtotals-all' ? 'Fruit Varianz (Grundgesamtheit)' : 'Fruit Summe',
			);
			createEditSession(wb).refreshPivotTable(s, pivot.name);
			const again = await loadXlsx(await saveXlsx(wb));
			expect(again.warnings).toEqual([]);
			expect(again.sheets[s]!.pivotTables![0]!.captions!.functions!.sum).toBe('Summe');
		},
	);

	it('writes the caption attributes Excel reads back', async () => {
		const wb = await loadXlsx(fixture('captions-de-columns'));
		const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
		createEditSession(wb).refreshPivotTable(s, wb.sheets[s]!.pivotTables![0]!.name);
		const root = await definition(await saveXlsx(wb));
		expect(root.getAttribute('dataCaption')).toBe('Werte');
		expect(root.getAttribute('rowHeaderCaption')).toBe('Zeilenbeschriftungen');
		expect(root.getAttribute('colHeaderCaption')).toBe('Spaltenbeschriftungen');
		expect(root.getAttribute('grandTotalCaption')).toBe('Gesamtergebnis');
	});

	it('does not add caption attributes to an English report', async () => {
		const wb = await loadXlsx(fixture('row-values-compact'));
		const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
		const pivot = wb.sheets[s]!.pivotTables![0]!;
		expect(pivot.captions).toBeUndefined();
		createEditSession(wb).refreshPivotTable(s, pivot.name);
		const root = await definition(await saveXlsx(wb));
		expect(root.getAttribute('dataCaption')).toBe('Values');
		for (const name of ['rowHeaderCaption', 'colHeaderCaption', 'grandTotalCaption'])
			expect(root.hasAttribute(name)).toBe(false);
	});

	it('names a new value field after the aggregate in the file language', async () => {
		const wb = await loadXlsx(fixture('captions-de-columns'));
		const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
		const edit = createEditSession(wb);
		const pivot = wb.sheets[s]!.pivotTables![0]!;
		edit.refreshPivotTable(s, pivot.name, {
			additionalValues: [
				{ field: 'Amount', aggregate: 'count' },
				{ field: 'Amount', aggregate: 'sum', caption: 'Mein Umsatz' },
			],
		});
		const next = wb.sheets[s]!.pivotTables![0]!;
		expect(next.additionalValues!.map((measure) => measure.caption)).toEqual([
			'Anzahl von Amount',
			'Mein Umsatz',
		]);
	});

	it('authors a report in the language of the editor', async () => {
		for (const [locale, rowLabels, total, sum] of [
			['de-DE', 'Zeilenbeschriftungen', 'Gesamtergebnis', 'Summe von Amount'],
			['es', 'Etiquetas de fila', 'Total general', 'Suma de Amount'],
			['fr-CA', 'Étiquettes de lignes', 'Total général', 'Somme de Amount'],
			['zh-CN', '行标签', '总计', '求和项:Amount'],
		] as const) {
			const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
			const edit = createEditSession(wb);
			edit.setRangeValues(0, { row: 0, col: 0 }, [
				['Category', 'Region', 'Amount'],
				['Fruit', 'East', 10],
				['Fruit', 'West', 20],
				['Veg', 'East', 5],
			] as never);
			edit.createPivotTable(
				1,
				{ row: 0, col: 0 },
				{
					source: 'Data!A1:C4',
					layout: 'compact',
					rowField: 'Category',
					secondRowField: 'Region',
					valueField: 'Amount',
					locale,
				},
			);
			const pivot = wb.sheets[1]!.pivotTables![0]!;
			expect(pivot.valueCaption).toBe(sum);
			const flat = pivotGrid(pivot).values.flat();
			expect(flat).toContain(rowLabels);
			expect(flat).toContain(total);
			expect(flat).not.toContain('Grand Total');
			const reopened = await loadXlsx(await saveXlsx(wb));
			expect(reopened.warnings).toEqual([]);
			const again = reopened.sheets[1]!.pivotTables![0]!;
			// The file keeps what its attributes and cells show.
			expect(again.captions).toMatchObject({ rowLabels, grandTotal: total });
			expect(pivotGrid(again).values).toEqual(pivotGrid(pivot).values);
			const root = await definition(await saveXlsx(wb));
			expect(root.getAttribute('rowHeaderCaption')).toBe(rowLabels);
			expect(root.getAttribute('grandTotalCaption')).toBe(total);
		}
	});

	it('keeps English for English and unsupported locales', () => {
		for (const locale of [undefined, 'en', 'en-GB', 'ja', 'pt-BR'])
			expect(localizedPivotCaptions(locale)).toBeUndefined();
		expect(pivotCaptions({})).toBe(ENGLISH_PIVOT_CAPTIONS);
	});

	it('has a complete table per language and distinct names for every aggregate', () => {
		for (const locale of ['de', 'es', 'fr', 'zh-CN']) {
			const caps = pivotCaptions({ captions: localizedPivotCaptions(locale)! });
			for (const key of ['rowLabels', 'columnLabels', 'grandTotal', 'values', 'all'] as const)
				expect(caps[key]).not.toBe(ENGLISH_PIVOT_CAPTIONS[key]);
			expect(caps.subtotal).toContain('{0}');
			expect(caps.measureTotal).toContain('{0}');
			expect(new Set(Object.values(caps.aggregates)).size).toBe(11);
			expect(new Set(Object.values(caps.functions)).size).toBeGreaterThanOrEqual(10);
		}
	});

	it('keeps the English subtotal labels in step with the subtotal table', () => {
		for (const [fn, entry] of Object.entries(PIVOT_SUBTOTALS))
			expect(ENGLISH_PIVOT_CAPTIONS.functions[fn as keyof typeof PIVOT_SUBTOTALS]).toBe(
				entry.label,
			);
	});
});

describe('item display names and the blank caption', () => {
	/** Renames East to Eastern the way Excel does: `item/@n`, and the cell shows the new name. */
	async function renamed(name: string) {
		const zip = await JSZip.loadAsync(fixture(name));
		const cache = parseXml(
			await zip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string'),
		).documentElement;
		const cacheFields = children(first(cache, 'cacheFields', NS.x)!, 'cacheField', NS.x);
		const field = cacheFields.findIndex((entry) =>
			children(first(entry, 'sharedItems', NS.x)!, 's', NS.x).some(
				(s) => s.getAttribute('v') === 'East',
			),
		);
		const at = children(first(cacheFields[field]!, 'sharedItems', NS.x)!, 's', NS.x).findIndex(
			(s) => s.getAttribute('v') === 'East',
		);
		const table = parseXml(await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string'));
		const pivotField = children(
			first(table.documentElement, 'pivotFields', NS.x)!,
			'pivotField',
			NS.x,
		)[field]!;
		children(first(pivotField, 'items', NS.x)!, 'item', NS.x)
			.find((item) => item.getAttribute('x') === String(at))!
			.setAttribute('n', 'Eastern');
		zip.file('xl/pivotTables/pivotTable1.xml', buildXml(table.documentElement));
		const strings = children(
			parseXml(await zip.file('xl/sharedStrings.xml')!.async('string')).documentElement,
			'si',
			NS.x,
		).map((si) => si.textContent);
		// The item and its subtotal rows (East Total) show the new name.
		let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		strings.forEach((text, index) => {
			if (text !== 'East' && !text?.startsWith('East ')) return;
			sheet = sheet.replace(
				new RegExp(`<c r="([A-Z]+\\d+)"((?: s="\\d+")?) t="s"><v>${index}</v></c>`, 'g'),
				`<c r="$1"$2 t="inlineStr"><is><t>Eastern${text.slice(4)}</t></is></c>`,
			);
		});
		zip.file('xl/worksheets/sheet1.xml', sheet);
		return zip.generateAsync({ type: 'uint8array' });
	}

	it.each(['rows3-compact-top', 'rows3-tabular'])(
		'shows an item under the display name the file gives it in %s',
		async (name) => {
			const wb = await loadXlsx(await renamed(name));
			expect(wb.warnings).toEqual([]);
			const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
			const pivot = wb.sheets[s]!.pivotTables![0]!;
			expect(pivot.itemNames).toEqual([{ field: 'Region', item: 'East', name: 'Eastern' }]);
			const flat = pivotGrid(pivot).values.flat();
			expect(flat).toContain('Eastern');
			expect(flat).not.toContain('East');
			createEditSession(wb).refreshPivotTable(s, pivot.name);
			const saved = await JSZip.loadAsync(await saveXlsx(wb));
			expect(await saved.file('xl/pivotTables/pivotTable1.xml')!.async('string')).toContain(
				'n="Eastern"',
			);
			const again = await loadXlsx(await saveXlsx(wb));
			expect(again.warnings).toEqual([]);
			expect(again.sheets[s]!.pivotTables![0]!.itemNames).toEqual(pivot.itemNames);
			expect(pivotGrid(again.sheets[s]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
		},
	);

	it.each(['captions-items-tabular', 'captions-items-compact-page'])(
		'reads the item names Excel wrote in %s and keeps them on refresh',
		async (name) => {
			const bytes = fixture(name);
			const wb = await loadXlsx(bytes);
			expect(wb.warnings).toEqual([]);
			const s = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
			const pivot = wb.sheets[s]!.pivotTables![0]!;
			expect(pivot.itemNames).toEqual([
				{ field: 'Category', item: 'Fruit', name: 'Obst' },
				{ field: 'Region', item: 'East', name: 'Osten' },
			]);
			const flat = pivotGrid(pivot).values.flat();
			expect(flat).toEqual(expect.arrayContaining(['Obst', 'Osten', 'Obst Total']));
			const original = await JSZip.loadAsync(bytes);
			const saved = await JSZip.loadAsync(await saveXlsx(wb));
			for (const part of Object.keys(original.files).filter((p) => p.startsWith('xl/pivot')))
				expect(await saved.file(part)?.async('string')).toBe(
					await original.file(part)!.async('string'),
				);
			const edit = createEditSession(wb);
			edit.setCellInput(
				wb.sheets.findIndex((sheet) => sheet.name === 'Data'),
				1,
				4,
				'100',
			);
			edit.refreshPivotTable(s, pivot.name);
			const after = pivotGrid(wb.sheets[s]!.pivotTables![0]!).values;
			expect(after.flat()).toContain('Obst Total');
			const again = await loadXlsx(await saveXlsx(wb));
			expect(again.warnings).toEqual([]);
			expect(pivotGrid(again.sheets[s]!.pivotTables![0]!).values).toEqual(after);
		},
	);

	it('labels a blank item in the language of the report', async () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		edit.setRangeValues(0, { row: 0, col: 0 }, [
			['Category', 'Amount'],
			['Fruit', 10],
			[null, 5],
		] as never);
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: 'Data!A1:B3', rowField: 'Category', valueField: 'Amount', locale: 'de' },
		);
		const pivot = wb.sheets[1]!.pivotTables![0]!;
		expect(pivotGrid(pivot).values.flat()).toContain('(leer)');
		const again = await loadXlsx(await saveXlsx(wb));
		expect(again.warnings).toEqual([]);
		expect(again.sheets[1]!.pivotTables![0]!.captions!.blank).toBe('(leer)');
		expect(pivotGrid(again.sheets[1]!.pivotTables![0]!).values.flat()).toContain('(leer)');
	});
});
