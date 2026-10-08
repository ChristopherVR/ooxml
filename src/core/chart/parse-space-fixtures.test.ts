// parseChartSpace against real chart parts: PowerPoint decks from the pptx fixture corpus (COM
// authored and generated) and the xlsx fixtures written by Excel and openpyxl.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { chartSourceValues } from './data-cache';
import { parseChartSpace } from './parse-space';

const root = import.meta.dirname;
const pptx = (name: string) => path.join(root, '../pptx/__tests__/fixtures/e2e', name);
const xlsx = (name: string) => path.join(root, '../xlsx/__fixtures__', name);

async function chartPart(file: string, part: string) {
	const zip = await JSZip.loadAsync(readFileSync(file));
	const xml = await zip.file(part)?.async('string');
	if (!xml) throw new Error(`${part} missing in ${file}`);
	return parseChartSpace(xml);
}

const issueCodes = (issues: { code: string; message: string }[]) =>
	issues.map((issue) => `${issue.code} ${issue.message.split(' ')[0]}`);

describe('parseChartSpace on chart-gallery.pptx (one chart kind per slide)', () => {
	const deck = pptx('chart-gallery.pptx');

	it('reads a clustered column chart: groups, series caches and axes', async () => {
		const { chartSpace, issues } = await chartPart(deck, 'ppt/charts/chart1.xml');
		expect(issues).toEqual([]);
		expect(chartSpace.title?.text).toBe('Clustered Bar');
		expect(chartSpace.autoTitleDeleted).toBe(false);
		expect(chartSpace.legend).toMatchObject({ position: 'b', overlay: false });
		expect(chartSpace.plotVisibleOnly).toBe(true);
		const [group] = chartSpace.plotArea.groups;
		expect(group).toMatchObject({
			kind: 'bar',
			element: 'barChart',
			is3D: false,
			barDirection: 'col',
			grouping: 'clustered',
			gapWidth: 150,
			overlap: -27,
			axisIds: [111111111, 222222222],
		});
		expect(group?.series).toHaveLength(2);
		const revenue = group?.series[0];
		expect(revenue).toMatchObject({ index: 0, order: 0 });
		expect(revenue?.tx).toMatchObject({ text: 'Revenue', reference: { formula: 'Sheet1!$B$1' } });
		expect(revenue?.categories).toMatchObject({ kind: 'strRef', formula: 'Sheet1!$A$2:$A$5' });
		expect(chartSourceValues(revenue?.categories)).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
		expect(revenue?.values?.cache).toMatchObject({ type: 'number', formatCode: 'General' });
		expect(chartSourceValues(revenue?.values)).toEqual([45, 62, 58, 71]);
		expect(revenue?.spPr?.fill?.kind).toBe('solid');
		expect(chartSpace.plotArea.axes.map((axis) => [axis.kind, axis.id, axis.position])).toEqual([
			['cat', 111111111, 'b'],
			['val', 222222222, 'l'],
		]);
		expect(chartSpace.plotArea.axes[1]).toMatchObject({
			crossAxisId: 111111111,
			deleted: false,
			scaling: { orientation: 'minMax' },
		});
	});

	it('reads every classic chart kind of the gallery', async () => {
		const kinds: Record<string, [string, number, string | undefined]> = {
			'chart2.xml': ['line', 2, 'Line (with trendline)'],
			'chart3.xml': ['area', 2, 'Area'],
			'chart4.xml': ['pie', 1, 'Pie'],
			'chart5.xml': ['doughnut', 1, 'Doughnut'],
			'chart6.xml': ['radar', 2, 'Radar'],
			'chart7.xml': ['scatter', 2, 'Scatter'],
			'chart8.xml': ['bubble', 2, 'Bubble'],
			'chart9.xml': ['bar', 3, 'Stacked Bar'],
			'chart16.xml': ['stock', 4, 'Stock'],
			'chart17.xml': ['surface', 3, 'Surface'],
		};
		for (const [part, [kind, series, title]] of Object.entries(kinds)) {
			const { chartSpace } = await chartPart(deck, `ppt/charts/${part}`);
			const group = chartSpace.plotArea.groups[0];
			expect([part, group?.kind, group?.series.length, chartSpace.title?.text]).toEqual([
				part,
				kind,
				series,
				title,
			]);
		}
	});

	it('reads kind-specific group settings', async () => {
		const group = async (part: string) =>
			(await chartPart(deck, `ppt/charts/${part}`)).chartSpace.plotArea.groups[0];
		expect(await group('chart5.xml')).toMatchObject({ holeSize: 50, varyColors: true });
		expect((await group('chart4.xml'))?.series[0]?.dataPoints).toHaveLength(4);
		expect(await group('chart6.xml')).toMatchObject({ radarStyle: 'marker' });
		const scatter = await group('chart7.xml');
		expect(scatter?.scatterStyle).toBe('marker');
		expect(scatter?.series[0]?.marker?.symbol).toBe('circle');
		expect(chartSourceValues(scatter?.series[0]?.xValues)).toEqual([1, 2, 3, 4]);
		expect(chartSourceValues(scatter?.series[0]?.yValues)).toEqual([45, 62, 58, 71]);
		const bubble = await group('chart8.xml');
		expect(chartSourceValues(bubble?.series[1]?.bubbleSizes)).toEqual([4, 5, 5, 7]);
		expect(await group('chart16.xml')).toMatchObject({ upDownBars: true, hiLowLines: true });
		expect(await group('chart17.xml')).toMatchObject({
			wireframe: false,
			axisIds: [111111111, 222222222, 333333333],
		});
	});

	it('reads a combination chart as two groups', async () => {
		const { chartSpace } = await chartPart(deck, 'ppt/charts/chart15.xml');
		const groups = chartSpace.plotArea.groups;
		expect(groups.map((group) => [group.kind, group.series.length])).toEqual([
			['bar', 1],
			['line', 2],
		]);
		expect(groups[1]?.marker).toBe(true);
	});

	it('reports trendlines and chartex parts instead of failing', async () => {
		const line = await chartPart(deck, 'ppt/charts/chart2.xml');
		expect(issueCodes(line.issues)).toEqual(['CHART_ELEMENT_NOT_MODELLED c:ser/c:trendline']);
		const chartex = await chartPart(deck, 'ppt/charts/chart11.xml');
		expect(chartex.issues.map((issue) => issue.code)).toEqual(['CHART_ROOT_UNEXPECTED']);
		expect(chartex.chartSpace.plotArea.groups).toEqual([]);
	});
});

describe('parseChartSpace on PowerPoint-authored charts', () => {
	it('reads bubble, scatter, sparse-category and pie charts (chart-data-fidelity.pptx)', async () => {
		const deck = pptx('chart-data-fidelity.pptx');
		const bubble = (await chartPart(deck, 'ppt/charts/chart1.xml')).chartSpace;
		expect(bubble).toMatchObject({ style: 2, roundedCorners: false, displayBlanksAs: 'gap' });
		expect(bubble.externalDataRelId).toBe('rId3');
		expect(bubble.chartExtLst).toContain('<c:extLst');
		const series = bubble.plotArea.groups[0]?.series ?? [];
		expect(series.map((entry) => entry.tx?.text)).toEqual(['Alpha Y', 'Beta Y', 'Gamma Y']);
		expect(series.map((entry) => chartSourceValues(entry.bubbleSizes))).toEqual([
			[4, 9, 2],
			[7, 3, 12],
			[2, 8, 5],
		]);
		expect(bubble.plotArea.axes.map((axis) => [axis.id, axis.crossBetween])).toEqual([
			[1995435263, 'midCat'],
			[782399823, 'midCat'],
		]);
		const scatter = (await chartPart(deck, 'ppt/charts/chart2.xml')).chartSpace;
		expect(scatter.plotArea.groups[0]?.scatterStyle).toBe('lineMarker');
		expect(scatter.plotArea.groups[0]?.series.map((entry) => entry.marker?.symbol)).toEqual([
			'none',
			'none',
		]);
		const column = (await chartPart(deck, 'ppt/charts/chart3.xml')).chartSpace;
		const units = column.plotArea.groups[0]?.series[0];
		expect(units?.categories?.cache?.pointCount).toBe(5);
		expect(chartSourceValues(units?.categories)).toEqual(['North', 'South', '', 'East', 'West']);
		expect(chartSourceValues(units?.values)).toEqual([12, 25, 7, 31, 18]);
		const pie = (await chartPart(deck, 'ppt/charts/chart4.xml')).chartSpace;
		const labels = pie.plotArea.groups[0]?.series[0]?.dataLabels;
		expect(labels).toMatchObject({
			showCategoryName: true,
			showPercent: true,
			showValue: false,
			separator: ', ',
		});
		expect(pie.plotArea.groups[0]?.dataLabels?.showPercent).toBe(false);
	});

	it('reads a top category axis over a reversed value axis (chart-top-axis.pptx)', async () => {
		const { chartSpace } = await chartPart(pptx('chart-top-axis.pptx'), 'ppt/charts/chart1.xml');
		expect(chartSpace.autoTitleDeleted).toBe(true);
		expect(chartSpace.legend).toBeUndefined();
		const [category, value] = chartSpace.plotArea.axes;
		expect(category).toMatchObject({ kind: 'cat', position: 't', majorTickMark: 'out' });
		expect(value).toMatchObject({ kind: 'val', scaling: { orientation: 'maxMin' } });
		expect(value?.numberFormat).toEqual({ formatCode: 'General', sourceLinked: true });
	});

	it('keeps the filtered-series extension raw (chart-filtered-series.pptx)', async () => {
		const { chartSpace } = await chartPart(
			pptx('chart-filtered-series.pptx'),
			'ppt/charts/chart1.xml',
		);
		const group = chartSpace.plotArea.groups[0];
		expect(group?.series.map((entry) => [entry.index, entry.tx?.text])).toEqual([
			[0, 'Series A'],
			[2, 'Series C'],
		]);
		expect(group?.extLst).toContain('filteredBarSeries');
		expect(group?.extLst).toContain('Series B');
		expect(chartSourceValues(group?.series[0]?.categories)).toEqual(['Cat1', 'Cat2', 'Cat4']);
	});

	it('reads rich title runs and stacked lines', async () => {
		const runs = (await chartPart(pptx('chart-title-runs.pptx'), 'ppt/charts/chart1.xml'))
			.chartSpace;
		expect(runs.title?.text).toBe('Sales Overview');
		expect(runs.title?.tx?.rich?.paragraphs[0]?.runs.length).toBeGreaterThan(1);
		const stacked = await chartPart(
			pptx('chart-stacked-line-markers.pptx'),
			'ppt/charts/chart1.xml',
		);
		const group = stacked.chartSpace.plotArea.groups[0];
		expect(group).toMatchObject({ kind: 'line', grouping: 'stacked', marker: true });
		expect(group?.series.map((entry) => [entry.index, entry.order, entry.smooth])).toEqual([
			[1, 0, false],
			[2, 1, false],
		]);
		expect(chartSourceValues(group?.series[0]?.categories)).toEqual([2011, 2012, 2013, 2014, 2015]);
		expect(stacked.chartSpace.displayBlanksAs).toBe('zero');
	});

	it('reads 3-D charts (three-d-charts.pptx, pie3d.pptx)', async () => {
		const deck = pptx('three-d-parity/three-d-charts.pptx');
		const bar = (await chartPart(deck, 'ppt/charts/chart1.xml')).chartSpace;
		expect(bar.plotArea.groups[0]).toMatchObject({
			element: 'bar3DChart',
			is3D: true,
			shape: 'box',
		});
		const surface = await chartPart(deck, 'ppt/charts/chart17.xml');
		expect(surface.chartSpace.plotArea.groups[0]).toMatchObject({
			kind: 'surface',
			is3D: true,
			wireframe: true,
		});
		expect(surface.chartSpace.plotArea.axes.map((axis) => axis.kind)).toEqual([
			'cat',
			'val',
			'ser',
		]);
		expect(surface.chartSpace.view3D).toEqual({ rotX: 15, rotY: 20, rightAngleAxes: false });
		expect(issueCodes(surface.issues)).toEqual([
			'CHART_ELEMENT_NOT_MODELLED c:chart/c:floor',
			'CHART_ELEMENT_NOT_MODELLED c:chart/c:sideWall',
			'CHART_ELEMENT_NOT_MODELLED c:chart/c:backWall',
			'CHART_ELEMENT_NOT_MODELLED c:surface3DChart/c:bandFmts',
		]);
		const pie = (await chartPart(pptx('pie3d.pptx'), 'ppt/charts/chart1.xml')).chartSpace;
		expect(pie.plotArea.groups[0]).toMatchObject({
			kind: 'pie',
			element: 'pie3DChart',
			is3D: true,
		});
		expect(pie.view3D).toMatchObject({ rotX: 30, rotY: 0 });
	});
});

describe('parseChartSpace on spreadsheet charts', () => {
	it('reads the Excel-authored chart (excel-features.xlsx)', async () => {
		const { chartSpace, issues } = await chartPart(
			xlsx('excel-features.xlsx'),
			'xl/charts/chart1.xml',
		);
		expect(chartSpace.title?.text).toBe('Sales by region');
		expect(chartSpace.style).toBe(2);
		const series = chartSpace.plotArea.groups[0]?.series ?? [];
		expect(series.map((entry) => [entry.tx?.text, entry.values?.formula])).toEqual([
			['Sales', 'Second!$B$2:$B$5'],
			['Cost', 'Second!$C$2:$C$5'],
		]);
		expect(chartSourceValues(series[0]?.values)).toEqual([100, 125, 150, 175]);
		expect(chartSpace.plotArea.axes.map((axis) => [axis.id, axis.majorGridlines])).toEqual([
			[1192221024, false],
			[1995301167, true],
		]);
		// Print settings are kept raw, so the Excel part has nothing unmodelled.
		expect(issues).toEqual([]);
		expect(chartSpace.printSettings).toContain('<c:pageMargins');
		expect(chartSpace.c14Style).toBe(102);
	});

	it('reads openpyxl charts, which carry references without caches', async () => {
		const file = xlsx('openpyxl-features.xlsx');
		const bar = (await chartPart(file, 'xl/charts/chart1.xml')).chartSpace;
		expect(bar.title?.text).toBe('Quarterly');
		expect(bar.legend?.position).toBe('r');
		expect(bar.plotArea.groups[0]?.axisIds).toEqual([10, 100]);
		const first = bar.plotArea.groups[0]?.series[0];
		expect(first?.tx).toEqual({ reference: { kind: 'strRef', formula: "'Table'!B1" } });
		expect(first?.categories).toEqual({ kind: 'numRef', formula: "'Table'!$A$2:$A$4" });
		expect(chartSourceValues(first?.values)).toEqual([]);
		const line = (await chartPart(file, 'xl/charts/chart2.xml')).chartSpace;
		expect(line.plotArea.groups[0]?.kind).toBe('line');
		const pie = (await chartPart(file, 'xl/charts/chart3.xml')).chartSpace;
		expect(pie.plotArea.groups[0]).toMatchObject({ kind: 'pie', firstSliceAngle: 0 });
	});
});
