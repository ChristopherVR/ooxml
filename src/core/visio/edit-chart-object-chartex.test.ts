import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import type { ChartExFamily } from '../chart/index';
import {
	createEditSession,
	createWorkbook,
	loadXlsx,
	saveXlsx,
	type CellValue,
} from '../xlsx/index';
import { buildChartEx } from '../xlsx/ui/dialogs/insert-chart-ex-model';
import { gridChartWorkbook, parseChartGrid } from '../xlsx/ui/dialogs/grid-chart';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { encodeBitmapEmf } from './emf-bitmap';
import { parseVsdx } from './parser';

const PREVIEW = encodeBitmapEmf(2, 2, new Uint8Array(16).fill(255));

const SAMPLE_ROWS: Record<ChartExFamily, readonly (readonly CellValue[])[]> = {
	waterfall: [
		['Item', 'Amount'],
		['Start', 100],
		['Sales', 60],
		['Returns', -20],
		['Costs', -45],
	],
	clusteredColumn: [
		['Observation', 'Score'],
		['A', 12],
		['B', 15],
		['C', 18],
		['D', 22],
		['E', 25],
		['F', 30],
	],
	pareto: [
		['Defect', 'Count'],
		['Scratch', 42],
		['Dent', 28],
		['Crack', 15],
		['Stain', 9],
	],
	boxWhisker: [
		['Group', 'Series 1', 'Series 2'],
		['North', 10, 14],
		['North', 14, 18],
		['North', 19, 21],
		['South', 8, 12],
		['South', 16, 20],
		['South', 22, 26],
	],
	treemap: [
		['Branch', 'Leaf', 'Size'],
		['Fruit', 'Apple', 30],
		['Fruit', 'Pear', 20],
		['Veg', 'Kale', 25],
		['Veg', 'Leek', 15],
	],
	sunburst: [
		['Branch', 'Leaf', 'Size'],
		['Fruit', 'Apple', 30],
		['Fruit', 'Pear', 20],
		['Veg', 'Kale', 25],
		['Veg', 'Leek', 15],
	],
	funnel: [
		['Stage', 'Leads'],
		['Visits', 1000],
		['Signups', 420],
		['Trials', 180],
		['Paid', 75],
	],
	regionMap: [
		['Country', 'Sales'],
		['United States', 120],
		['Canada', 85],
		['France', 64],
		['Australia', 52],
	],
};

const EXPECTED_VIEW_TYPE: Record<ChartExFamily, string> = {
	waterfall: 'waterfall',
	clusteredColumn: 'histogram',
	pareto: 'histogram',
	boxWhisker: 'boxWhisker',
	treemap: 'treemap',
	sunburst: 'sunburst',
	funnel: 'funnel',
	regionMap: 'regionMap',
};

async function buildEmbeddedChartExWorkbook(
	family: ChartExFamily,
	title: string,
): Promise<Uint8Array> {
	const rows = SAMPLE_ROWS[family];
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setRangeValues(
		0,
		{ row: 0, col: 0 },
		rows.map((row) => [...row]),
	);
	const range = {
		start: { row: 0, col: 0 },
		end: { row: rows.length - 1, col: rows[0]!.length - 1 },
	};
	const built = buildChartEx(workbook, 0, range, family, { title });
	session.addChartEx(0, built);
	return saveXlsx(workbook);
}

async function embedWorkbookInVsdx(workbook: Uint8Array): Promise<Uint8Array> {
	const result = await editVsdx(await createVsdx(), [
		{
			type: 'insert-chart-object',
			pageId: '0',
			shapeId: '1',
			x: 4.25,
			y: 5.5,
			width: 6.4,
			height: 4,
			workbook,
			preview: PREVIEW,
		},
	]);
	return result.bytes;
}

describe('embedded ChartEx workbooks in Visio', () => {
	const families: readonly ChartExFamily[] = [
		'waterfall',
		'clusteredColumn',
		'pareto',
		'boxWhisker',
		'treemap',
		'sunburst',
		'funnel',
		'regionMap',
	];

	it.each(families)(
		'parses an embedded %s ChartEx workbook into VisioEmbeddedObject',
		async (family) => {
			const title = `Visio ${family}`;
			const vsdx = await embedWorkbookInVsdx(await buildEmbeddedChartExWorkbook(family, title));
			const doc = await parseVsdx(vsdx);
			const shape = doc.pages[0]!.shapes[0]!;
			expect(doc.diagnostics.map((d) => d.code)).not.toContain('unsupported-foreign-object');
			expect(shape.embeddedObject).toBeDefined();
			expect(shape.embeddedObject!.chartEx).toBe(true);
			expect(shape.embeddedObject!.chart.type).toBe(EXPECTED_VIEW_TYPE[family]);
			expect(shape.embeddedObject!.chart.supported).toBe(true);
			expect(shape.embeddedObject!.chart.title).toBe(title);
			expect(shape.embeddedObject!.chart.series.length).toBeGreaterThan(0);
			if (family === 'pareto') {
				expect(shape.embeddedObject!.chart.secondaryAxis?.percent).toBe(true);
				expect(shape.embeddedObject!.chart.extraGroups?.[0]?.type).toBe('line');
			}
			if (family === 'treemap' || family === 'sunburst')
				expect(shape.embeddedObject!.chart.hierarchy?.branches).toEqual(['Fruit', 'Veg']);
			if (family === 'waterfall')
				expect(shape.embeddedObject!.chart.waterfall?.steps.length).toBe(4);
			expect(() => structuredClone(shape.embeddedObject)).not.toThrow();
		},
	);

	it('parses an Excel-authored ChartEx workbook embedded in a Visio page', async () => {
		const fixture = new Uint8Array(
			readFileSync(
				path.join(import.meta.dirname, '..', 'xlsx', '__fixtures__', 'excel-chart-families.xlsx'),
			),
		);
		const book = await loadXlsx(fixture);
		for (const ws of book.sheets)
			ws.drawings = ws.drawings.filter(
				(d) =>
					d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall',
			);
		const vsdx = await embedWorkbookInVsdx(await saveXlsx(book));
		const doc = await parseVsdx(vsdx);
		const embedded = doc.pages[0]!.shapes[0]!.embeddedObject!;
		expect(embedded.chartEx).toBe(true);
		expect(embedded.chart.type).toBe('waterfall');
		expect(embedded.chart.title).toBe('Cash flow');
		expect(embedded.chart.waterfall?.steps.map((s) => s.endValue)).toEqual([
			100, 160, 140, 95, 110, 110,
		]);
	});

	it('falls back to standalone xl/charts/chartEx*.xml when the workbook has no worksheet drawing', async () => {
		const fullWorkbook = await buildEmbeddedChartExWorkbook('funnel', 'Standalone Funnel');
		const sourceZip = await JSZip.loadAsync(fullWorkbook);
		const chartExXml = await sourceZip.file('xl/charts/chartEx1.xml')!.async('string');
		const standalone = new JSZip();
		standalone.file('xl/charts/chartEx1.xml', chartExXml);
		const themeFile = sourceZip.file('xl/theme/theme1.xml');
		if (themeFile) standalone.file('xl/theme/theme1.xml', await themeFile.async('string'));
		const vsdx = await embedWorkbookInVsdx(await standalone.generateAsync({ type: 'uint8array' }));
		const doc = await parseVsdx(vsdx);
		const embedded = doc.pages[0]!.shapes[0]!.embeddedObject!;
		expect(embedded.chartEx).toBe(true);
		expect(embedded.chart.type).toBe('funnel');
		expect(embedded.chart.title).toBe('Standalone Funnel');
		expect(embedded.chart.series[0]?.values).toEqual([1000, 420, 180, 75]);
	});

	it('leaves chartEx unset for classic c:chartSpace embedded workbooks', async () => {
		const classic = (await gridChartWorkbook(
			parseChartGrid('Month\tSales\nJan\t10\nFeb\t20'),
			'column',
			{ title: 'Classic' },
		))!;
		const vsdx = await embedWorkbookInVsdx(classic);
		const embedded = (await parseVsdx(vsdx)).pages[0]!.shapes[0]!.embeddedObject!;
		expect(embedded.chartEx).toBeUndefined();
		expect(embedded.chart.type).toBe('column');
		expect(embedded.chart.title).toBe('Classic');
	});
});
