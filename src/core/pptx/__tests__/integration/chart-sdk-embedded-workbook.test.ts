/**
 * `addChart(..., { embedWorkbook: true })`: a new classic chart is saved with its data in an
 * embedded workbook the chart reads through cell references, the parts PowerPoint writes for
 * Insert > Chart (chart part, `package` relationship, `.xlsx`, content types). The reloaded chart
 * is a normal chart with an embedded workbook, whose edits go through the loaded-chart path.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import type { ChartInput } from '../../core/builders/sdk/types';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxChartType, PptxSlide } from '../../core/types';

const WORKBOOK = 'ppt/embeddings/Microsoft_Excel_Worksheet.xlsx';
const RELS = 'ppt/charts/_rels/chart1.xml.rels';
const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const COLOR_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const chartOf = (slide: PptxSlide | undefined) =>
	slide?.elements.find((item): item is ChartPptxElement => item.type === 'chart');

async function build(
	type: PptxChartType,
	input: ChartInput,
	edit?: (chart: PptxChartData) => void,
): Promise<{ handler: PptxHandler; slides: PptxSlide[] }> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addChart(type, { ...input, embedWorkbook: true }, { x: 50, y: 50, width: 500, height: 300 })
		.build();
	edit?.(chartOf(slide)!.chartData!);
	data.slides.push(slide);
	return { handler, slides: data.slides };
}

interface SavedChart {
	bytes: Uint8Array;
	xml: string;
	rels: string;
	contentTypes: string;
	/** Cell text by reference (`B2`), from the workbook's only sheet. */
	cells: Map<string, string>;
}

async function cellsOf(zip: JSZip): Promise<Map<string, string>> {
	const workbook = await JSZip.loadAsync(await zip.file(WORKBOOK)!.async('uint8array'));
	const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
	const cells = new Map<string, string>();
	for (const match of sheet.matchAll(/<c r="([A-Z]+\d+)"[^>]*>(.*?)<\/c>/gsu))
		cells.set(match[1]!, /<(?:v|t)[^>]*>([^<]*)</u.exec(match[2]!)?.[1] ?? '');
	return cells;
}

async function saved(handler: PptxHandler, slides: PptxSlide[]): Promise<SavedChart> {
	const bytes = await handler.save(slides);
	const zip = await JSZip.loadAsync(bytes);
	return {
		bytes,
		xml: await zip.file('ppt/charts/chart1.xml')!.async('string'),
		rels: await zip.file(RELS)!.async('string'),
		contentTypes: await zip.file('[Content_Types].xml')!.async('string'),
		cells: await cellsOf(zip),
	};
}

const formulas = (xml: string) => [...xml.matchAll(/<c:f>([^<]*)<\/c:f>/gu)].map((m) => m[1]);
const relIdOf = (xml: string) => /<c:externalData r:id="([^"]+)"/u.exec(xml)?.[1];

function expectWorkbookPackage(result: SavedChart): void {
	const id = relIdOf(result.xml);
	expect(id).toBeDefined();
	expect(result.xml).toContain('<c:autoUpdate val="0"/>');
	expect(result.rels).toContain(
		`Id="${id}" Type="${PACKAGE_REL}" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"`,
	);
	expect(result.contentTypes).toContain(
		`PartName="/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx" ContentType="${XLSX}"`,
	);
}

const CATEGORIES = ['Q1', 'Q2', 'Q3', 'Q4'];

describe('new charts with an embedded workbook', () => {
	it.each(['bar', 'line', 'pie'] as const)(
		'writes a %s chart in the standard layout',
		async (type) => {
			const series =
				type === 'pie'
					? [{ name: 'Share', values: [4, null, 2, 1] }]
					: [
							{ name: 'North', values: [4, null, 2, 1] },
							{ name: 'South', values: [3, 5, 7, 9] },
						];
			const { handler, slides } = await build(type, { categories: CATEGORIES, series });
			const result = await saved(handler, slides);
			expectWorkbookPackage(result);
			expect(formulas(result.xml)).toStrictEqual(
				series.flatMap((_, index) => {
					const col = String.fromCharCode(66 + index);
					return [`Sheet1!$${col}$1`, 'Sheet1!$A$2:$A$5', `Sheet1!$${col}$2:$${col}$5`];
				}),
			);
			// Caches stay with the references; the blank point is left out of the value cache.
			expect(result.xml).toContain('<c:strRef><c:f>Sheet1!$A$2:$A$5</c:f><c:strCache>');
			const firstValues = /<c:val><c:numRef>.*?<\/c:numRef><\/c:val>/su.exec(result.xml)![0];
			expect(firstValues).toContain('<c:ptCount val="4"/>');
			expect(firstValues).not.toContain('<c:pt idx="1">');
			expect(result.cells.get('A2')).toBe('Q1');
			expect(result.cells.get('A5')).toBe('Q4');
			expect(result.cells.get('B1')).toBe(series[0]!.name);
			expect(result.cells.get('B2')).toBe('4');
			expect(result.cells.has('B3')).toBe(false);
			expect(result.cells.get('B5')).toBe('1');
			if (type !== 'pie') expect(result.cells.get('C5')).toBe('9');
			expect(result.cells.has('A1')).toBe(false);
		},
	);

	it('writes scatter X values to their own columns', async () => {
		const { handler, slides } = await build('scatter', {
			categories: CATEGORIES,
			series: [{ name: 'Fit', values: [1, 4, null, 16], xValues: [1, 2, 3, 4] }],
		});
		const result = await saved(handler, slides);
		expectWorkbookPackage(result);
		expect(formulas(result.xml)).toStrictEqual([
			'Sheet1!$B$1',
			'Sheet1!$A$2:$A$5',
			'Sheet1!$B$2:$B$5',
		]);
		expect(result.xml).toMatch(/<c:xVal><c:numRef><c:f>Sheet1!\$A\$2:\$A\$5<\/c:f><c:numCache>/u);
		expect(result.cells.get('A1')).toBe('Fit X');
		expect(result.cells.get('B1')).toBe('Fit');
		expect(['A2', 'A3', 'A4', 'A5'].map((ref) => result.cells.get(ref))).toStrictEqual([
			'1',
			'2',
			'3',
			'4',
		]);
		expect(result.cells.get('B5')).toBe('16');
		expect(result.cells.has('B4')).toBe(false);
	});

	it('writes bubble sizes next to the values and keeps the colour-style relationship', async () => {
		const { handler, slides } = await build(
			'bubble',
			{
				categories: CATEGORIES,
				series: [
					{ name: 'B', values: [1, 2, 3, 4], xValues: [5, 6, 7, 8], bubbleSizes: [9, 8, 7, 6] },
				],
			},
			(chart) => {
				chart.colorPalette = ['#112233', '#445566'];
			},
		);
		const result = await saved(handler, slides);
		expectWorkbookPackage(result);
		expect(formulas(result.xml)).toStrictEqual([
			'Sheet1!$B$1',
			'Sheet1!$A$2:$A$5',
			'Sheet1!$B$2:$B$5',
			'Sheet1!$C$2:$C$5',
		]);
		expect(result.xml).toMatch(/<c:bubbleSize><c:numRef><c:f>Sheet1!\$C\$2:\$C\$5<\/c:f>/u);
		expect(result.cells.get('C1')).toBe('B Size');
		expect(result.cells.get('C2')).toBe('9');
		expect(result.cells.get('A5')).toBe('8');
		// One relationships part naming both the workbook and the colour style, with distinct ids.
		const ids = [...result.rels.matchAll(/Id="([^"]+)"/gu)].map((m) => m[1]);
		expect(ids).toHaveLength(2);
		expect(new Set(ids).size).toBe(2);
		expect(result.rels).toContain(`Type="${COLOR_REL}" Target="colors1.xml"`);
		expect(result.contentTypes).toContain('/ppt/charts/colors1.xml');
	});

	it('reloads as a chart with an embedded workbook whose edits update it', async () => {
		const { handler, slides } = await build('bar', {
			categories: CATEGORIES,
			series: [{ name: 'North', values: [4, null, 2, 1] }],
		});
		const first = await saved(handler, slides);
		const reloadHandler = new PptxHandler();
		const reloaded = await reloadHandler.load(first.bytes.buffer as ArrayBuffer);
		const chart = chartOf(reloaded.slides[0])!.chartData!;
		expect(chart.embedWorkbook).toBeUndefined();
		expect(chart.externalData?.targetPath).toContain('Microsoft_Excel_Worksheet.xlsx');
		expect(chart.categories).toStrictEqual(CATEGORIES);
		expect(chart.series[0]!.values).toStrictEqual([4, 0, 2, 1]);
		expect(chart.series[0]!.blanks).toStrictEqual([false, true, false, false]);

		chart.series[0]!.values = [4, 0, 42, 1];
		const second = await saved(reloadHandler, reloaded.slides);
		expect(second.cells.get('B4')).toBe('42');
		expect(second.xml).toMatch(/<c:pt idx="2"><c:v>42<\/c:v><\/c:pt>/u);
		expect(relIdOf(second.xml)).toBe(relIdOf(first.xml));
		const again = await new PptxHandler().load(second.bytes.buffer as ArrayBuffer);
		expect(chartOf(again.slides[0])!.chartData!.series[0]!.values[2]).toBe(42);
	});

	it('names two charts on one slide after distinct workbooks', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const input = { categories: ['A'], series: [{ name: 'S', values: [1] }], embedWorkbook: true };
		const slide = createSlide('Blank')
			.addChart('bar', input, { x: 0, y: 0, width: 200, height: 200 })
			.addChart('line', input, { x: 200, y: 0, width: 200, height: 200 })
			.build();
		data.slides.push(slide);
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const rels = await zip.file('ppt/charts/_rels/chart2.xml.rels')!.async('string');
		expect(rels).toContain('Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"');
		expect(zip.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')).not.toBeNull();
		expect(zip.file(WORKBOOK)).not.toBeNull();
	});

	it('refuses a chart the shared chart core cannot write instead of dropping the workbook', async () => {
		const { handler, slides } = await build('stock', {
			categories: CATEGORIES,
			series: [
				{ name: 'High', values: [3, 4, 5, 6] },
				{ name: 'Low', values: [1, 2, 3, 4] },
			],
		});
		await expect(handler.save(slides)).rejects.toThrow(
			/asks for an embedded workbook \(embedWorkbook\), but the shared chart core cannot write it \(refused: A stock chart requires three or four price series\)/u,
		);
	});

	it('refuses an extended chart type when the chart is created', async () => {
		const { createSlide } = await PresentationBuilder.create();
		expect(() =>
			createSlide('Blank').addChart('funnel', {
				categories: ['A'],
				series: [{ name: 'S', values: [1] }],
				embedWorkbook: true,
			}),
		).toThrow(/embedWorkbook applies to classic charts/u);
	});
});
