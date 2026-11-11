import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { ChartPlotGroup, ChartSpace } from '../model';
import { categorySource, numberSource } from '../edit/series-data';
import { newChartSpace } from '../edit/new-chart';
import { parseChartSpace } from '../parse-space';
import { newChartWorkbookLayout } from './new-layout';
import {
	CHART_WORKBOOK_RELATIONSHIP_TYPE,
	stageChartPartWithWorkbook,
	writeChartPartWithWorkbook,
} from './new-part';

/** A literal chart: one `kind` group whose series read `categories` and `values`. */
function literalChart(
	group: Partial<ChartPlotGroup>,
	series: { name: string; values: (number | null)[]; x?: number[]; sizes?: number[] }[],
	categories: (string | number)[] = ['A', 'B', 'C'],
): ChartSpace {
	return {
		plotArea: {
			layout: {},
			axes: [],
			groups: [
				{
					kind: 'bar',
					element: 'barChart',
					is3D: false,
					axisIds: [],
					...group,
					series: series.map((entry, index) => ({
						index,
						order: index,
						tx: { value: entry.name, text: entry.name },
						dataPoints: [],
						...(entry.x
							? {
									xValues: numberSource(entry.x, undefined),
									yValues: numberSource(entry.values, undefined),
								}
							: {
									categories: categorySource(categories, undefined),
									values: numberSource(entry.values, undefined),
								}),
						...(entry.sizes ? { bubbleSizes: numberSource(entry.sizes, undefined) } : {}),
					})),
				} as ChartPlotGroup,
			],
		},
	};
}

describe('newChartWorkbookLayout', () => {
	it('lays shared categories out in the standard layout, blanks as empty cells', () => {
		const space = literalChart({}, [
			{ name: 'North', values: [1, null, 3] },
			{ name: 'South', values: [4, 5, 6] },
		]);
		const { chartSpace, rows } = newChartWorkbookLayout(space);
		expect(rows).toEqual([
			[null, 'North', 'South'],
			['A', 1, 4],
			['B', null, 5],
			['C', 3, 6],
		]);
		const second = chartSpace.plotArea.groups[0]!.series[1]!;
		expect(second.tx?.reference?.formula).toBe('Sheet1!$C$1');
		expect(second.categories).toMatchObject({ kind: 'strRef', formula: 'Sheet1!$A$2:$A$4' });
		expect(second.values).toMatchObject({ kind: 'numRef', formula: 'Sheet1!$C$2:$C$4' });
		expect(chartSpace.plotArea.groups[0]!.series[0]!.values!.cache!.points).toHaveLength(2);
		expect(space.plotArea.groups[0]!.series[0]!.values!.kind).toBe('numLit');
	});

	it('keeps numeric categories as number cells', () => {
		const { chartSpace, rows } = newChartWorkbookLayout(
			literalChart({}, [{ name: 'S', values: [1, 2] }], [45000, 45001]),
		);
		expect(rows.slice(1).map((row) => row[0])).toEqual([45000, 45001]);
		expect(chartSpace.plotArea.groups[0]!.series[0]!.categories!.kind).toBe('numRef');
	});

	it('gives scatter and bubble vectors their own columns', () => {
		const space = literalChart({ kind: 'bubble', element: 'bubbleChart' }, [
			{ name: 'B', values: [1, 2], x: [5, 6], sizes: [9, 8] },
		]);
		const { chartSpace, rows } = newChartWorkbookLayout(space);
		expect(rows).toEqual([
			['B X', 'B', 'B Size'],
			[5, 1, 9],
			[6, 2, 8],
		]);
		const series = chartSpace.plotArea.groups[0]!.series[0]!;
		expect(series.xValues?.formula).toBe('Sheet1!$A$2:$A$3');
		expect(series.bubbleSizes?.formula).toBe('Sheet1!$C$2:$C$3');
	});
});

describe('stageChartPartWithWorkbook', () => {
	it('writes the part, merged relationships and a workbook from the given rows', async () => {
		const zip = new JSZip();
		zip.file('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx', new Uint8Array([1]));
		const { chartSpace, rows } = newChartWorkbookLayout(
			literalChart({}, [{ name: 'S', values: [7, null, 9] }]),
		);
		const parts = stageChartPartWithWorkbook(zip, chartSpace, {
			chartFolder: 'ppt/charts/',
			embeddingsFolder: 'ppt/embeddings/',
			partName: 'ppt/charts/chart4.xml',
			relationships: new Map([
				['rId1', { type: 'urn:colors', target: 'colors4.xml', mode: 'Internal' as const }],
			]),
			rows,
		});
		expect(parts).toEqual({
			partName: 'ppt/charts/chart4.xml',
			workbookPartName: 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',
			workbookRelationshipId: 'rId2',
			contentTypes: [
				[
					'ppt/charts/chart4.xml',
					'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
				],
				[
					'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',
					'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				],
			],
		});
		const xml = await zip.file('ppt/charts/chart4.xml')!.async('string');
		expect(parseChartSpace(xml).chartSpace.externalDataRelId).toBe('rId2');
		const rels = await zip.file('ppt/charts/_rels/chart4.xml.rels')!.async('string');
		expect(rels).toContain(
			`<Relationship Id="rId2" Type="${CHART_WORKBOOK_RELATIONSHIP_TYPE}" Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/>`,
		);
		expect(rels).toContain('<Relationship Id="rId1" Type="urn:colors" Target="colors4.xml"/>');
		const workbook = await JSZip.loadAsync(
			await zip.file(parts.workbookPartName)!.async('uint8array'),
		);
		const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(sheet).toContain('<c r="B2"><v>7</v></c>');
		expect(sheet).not.toContain('r="B3"');
	});

	it('adds the content types and picks a free chart part name', async () => {
		const zip = new JSZip();
		zip.file(
			'[Content_Types].xml',
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
		);
		zip.file('word/charts/chart1.xml', '');
		const parts = await writeChartPartWithWorkbook(zip, newChartSpace(), {
			chartFolder: 'word/charts/',
			embeddingsFolder: 'word/embeddings/',
		});
		expect(parts.partName).toBe('word/charts/chart2.xml');
		expect(parts.workbookRelationshipId).toBe('rId1');
		const types = await zip.file('[Content_Types].xml')!.async('string');
		expect(types).toContain('PartName="/word/charts/chart2.xml"');
		expect(types).toContain('PartName="/word/embeddings/Microsoft_Excel_Worksheet.xlsx"');
	});
});
