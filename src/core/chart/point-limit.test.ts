import { describe, expect, it } from 'vitest';
import { chartCacheValues } from './data-cache';
import { addChartExSeries } from './edit/chartex-structure';
import { chartExLevelValues, chartExSpaceDataSheet } from './edit/chartex-sheet';
import { chartSpaceDataSheet } from './edit/data-sheet';
import { parseChartExSpace } from './ex/parse';
import { writeChartExSpace } from './ex/write';
import { writeChartExSpaceOverSource } from './ex/keep-source';
import { chartIssuesLostOnWrite, writeChartSpaceOverSource } from './keep-source';
import { parseChartSpace } from './parse-space';
import { CHART_POINTS_LIMITED, MAX_CHART_BLANK_POINTS } from './point-limit';
import { chartExSummary } from './render/summary-from-chartex';
import { renderChartSummarySvg } from './render/render-space';
import { chartSummaryFromSpace } from './render/summary-from-space';
import { writeChartSpace } from './write-space';

const HUGE = 2_000_000_000;
const C =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const CX =
	'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const numCache = (inner: string) =>
	`<c:numRef><c:f>S!$B$2:$B$9</c:f><c:numCache><c:formatCode>General</c:formatCode>${inner}</c:numCache></c:numRef>`;
const series = (index: number, cache: string, categories = '') =>
	`<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${categories}<c:val>${numCache(cache)}</c:val></c:ser>`;
const barChart = (body: string) =>
	`<c:chartSpace ${C}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${body}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
const chartEx = (levels: string, layout = 'clusteredColumn') =>
	`<cx:chartSpace ${CX}><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>S!$A$1:$A$3</cx:f>${levels}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;

const firstValues = (xml: string) =>
	parseChartSpace(xml).chartSpace.plotArea.groups[0]!.series[0]!.values!;
const codes = (issues: readonly { code: string }[]) => issues.map((issue) => issue.code);

describe('classic chart caches', () => {
	it('keeps sparse caches, their holes and their declared count', () => {
		const xml = barChart(
			series(
				0,
				'<c:ptCount val="6"/><c:pt idx="1"><c:v>10</c:v></c:pt><c:pt idx="4"><c:v>x</c:v></c:pt>',
			),
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		const values = chartSpace.plotArea.groups[0]!.series[0]!.values!;
		expect(issues).toEqual([]);
		expect(values.cache).toMatchObject({ pointCount: 6 });
		expect(values.cache!.points.map((point) => point.index)).toEqual([1, 4]);
		expect(chartCacheValues(values.cache)).toEqual(['', 10, '', '', 'x', '']);
		const written = writeChartSpace(chartSpace);
		expect(written).toContain('<c:ptCount val="6"/>');
		expect(written).toContain('<c:pt idx="4">');
	});

	it('keeps a cache of real points however long it is', () => {
		const count = MAX_CHART_BLANK_POINTS + 10;
		const points = Array.from({ length: count }, (_, index) => ({ index, value: '1' }));
		expect(chartCacheValues({ type: 'number', pointCount: count, points })).toHaveLength(count);
	});

	it('does not expand a declared count or an index from a hand-built cache', () => {
		const counted = chartCacheValues({ type: 'number', pointCount: HUGE, points: [] });
		expect(counted.length).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS);
		const indexed = chartCacheValues({ type: 'string', points: [{ index: HUGE, value: 'x' }] });
		expect(indexed.length).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS + 1);
	});

	it('reports a declared count past the limit and bounds what it expands to', () => {
		const xml = barChart(series(0, `<c:ptCount val="${HUGE}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`));
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(codes(issues)).toContain(CHART_POINTS_LIMITED);
		expect(chartIssuesLostOnWrite(issues).map((issue) => issue.code)).toContain(
			CHART_POINTS_LIMITED,
		);
		const summary = chartSummaryFromSpace(chartSpace)!;
		expect(summary.series[0]!.values[0]).toBe(1);
		expect(summary.series[0]!.values.length).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS + 1);
	});

	it('drops a point whose index is past the limit and reports it', () => {
		const xml = barChart(
			series(
				0,
				`<c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="${HUGE}"><c:v>2</c:v></c:pt>`,
			),
		);
		const { issues } = parseChartSpace(xml);
		expect(codes(issues)).toContain(CHART_POINTS_LIMITED);
		expect(chartCacheValues(firstValues(xml).cache)).toEqual([1, '']);
	});

	it('shares the allowance among the caches of a part, so series cannot multiply it', () => {
		const count = 2000;
		const xml = barChart(
			Array.from({ length: count }, (_, index) =>
				series(
					index,
					`<c:ptCount val="${MAX_CHART_BLANK_POINTS}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`,
				),
			).join(''),
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(codes(issues)).toContain(CHART_POINTS_LIMITED);
		const sheet = chartSpaceDataSheet(chartSpace);
		expect(sheet.series).toHaveLength(count);
		expect(sheet.categories.length * count).toBeLessThanOrEqual(2 * MAX_CHART_BLANK_POINTS + count);
	});

	it('bounds the count a multi-level category cache declares for every level', () => {
		const levels = '<c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt></c:lvl>'.repeat(50);
		const categories = `<c:cat><c:multiLvlStrRef><c:f>S!$A$1</c:f><c:multiLvlStrCache><c:ptCount val="${HUGE}"/>${levels}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
		const { chartSpace, issues } = parseChartSpace(
			barChart(series(0, '<c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>', categories)),
		);
		const source = chartSpace.plotArea.groups[0]!.series[0]!.categories!;
		expect(codes(issues)).toContain(CHART_POINTS_LIMITED);
		expect(source.levelPointCount! * 50).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS + 50);
	});

	it('keeps a legitimate multi-level count as written', () => {
		const categories =
			'<c:cat><c:multiLvlStrRef><c:f>S!$A$1</c:f><c:multiLvlStrCache><c:ptCount val="4"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>';
		const { chartSpace, issues } = parseChartSpace(
			barChart(series(0, '<c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>', categories)),
		);
		expect(issues).toEqual([]);
		expect(chartSpace.plotArea.groups[0]!.series[0]!.categories!.levelPointCount).toBe(4);
	});
});

describe('saving a limited part', () => {
	it('keeps the classic source bytes and refuses to rewrite them from the shortened model', () => {
		const xml = barChart(series(0, `<c:ptCount val="${HUGE}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`));
		const { chartSpace } = parseChartSpace(xml);
		expect(writeChartSpaceOverSource(xml, chartSpace)).toBe(xml);
		expect(() => writeChartSpaceOverSource(xml, { ...chartSpace, roundedCorners: true })).toThrow();
	});

	it('keeps the ChartEx source bytes and refuses to rewrite them', () => {
		const xml = chartEx(`<cx:lvl ptCount="${HUGE}"><cx:pt idx="0">1</cx:pt></cx:lvl>`);
		const { chartExSpace } = parseChartExSpace(xml);
		expect(writeChartExSpaceOverSource(xml, chartExSpace)).toBe(xml);
		expect(() =>
			writeChartExSpaceOverSource(xml, { ...chartExSpace, title: { tx: { value: 'New' } } }),
		).toThrow();
	});
});

describe('rendering', () => {
	it('keeps a radar with very many categories to a bounded drawing', () => {
		const radar = (categories: number) =>
			barChart(series(0, `<c:ptCount val="${categories}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`))
				.replace(
					'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>',
					'<c:radarChart><c:radarStyle val="marker"/>',
				)
				.replace('</c:barChart>', '</c:radarChart>');
		const svg = (categories: number) =>
			renderChartSummarySvg(chartSummaryFromSpace(parseChartSpace(radar(categories)).chartSpace)!, {
				width: 300,
				height: 200,
			});
		expect(svg(100_000).length).toBeLessThan(svg(360).length + 3_000_000);
	});
});

describe('ChartEx levels', () => {
	it('keeps sparse levels and writes them back as authored', () => {
		const xml = chartEx(
			'<cx:lvl ptCount="5"><cx:pt idx="1">2</cx:pt><cx:pt idx="3">4</cx:pt></cx:lvl>',
		);
		const { chartExSpace, issues } = parseChartExSpace(xml);
		expect(issues).toEqual([]);
		const level = chartExSpace.chartData.data[0]!.dimensions[0]!.levels[0]!;
		expect(chartExLevelValues(level)).toEqual(['', '2', '', '4', '']);
		expect(writeChartExSpace(chartExSpace)).toContain('ptCount="5"');
		expect(chartExSummary(xml)).not.toHaveProperty('pointsLimited');
		expect(chartExSummary(xml)!.summary!.series[0]!.values).toEqual([null, 2, null, 4, null]);
	});

	it('reports a huge declared count and stays within the limit', () => {
		const xml = chartEx(`<cx:lvl ptCount="${HUGE}"><cx:pt idx="0">1</cx:pt></cx:lvl>`);
		const result = chartExSummary(xml)!;
		expect(result.pointsLimited).toBe(true);
		expect(result.summary!.series[0]!.values[0]).toBe(1);
		expect(result.summary!.series[0]!.values.length).toBeLessThanOrEqual(
			MAX_CHART_BLANK_POINTS + 1,
		);
		expect(codes(parseChartExSpace(xml).issues)).toContain(CHART_POINTS_LIMITED);
	});

	it('drops a point whose index is past the limit', () => {
		const xml = chartEx(
			`<cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt><cx:pt idx="${HUGE}">2</cx:pt></cx:lvl>`,
		);
		const result = chartExSummary(xml)!;
		expect(result.pointsLimited).toBe(true);
		expect(result.summary!.series[0]!.values).toEqual([1]);
	});

	it('bounds the data sheet and a new series of an oversized level', () => {
		const { chartExSpace } = parseChartExSpace(
			chartEx(`<cx:lvl ptCount="${HUGE}"><cx:pt idx="0">1</cx:pt></cx:lvl>`),
		);
		const sheet = chartExSpaceDataSheet(chartExSpace);
		expect(sheet.categories.length).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS + 1);
		const added = addChartExSeries(chartExSpace, 'Copy');
		const level = added.chartData.data.at(-1)!.dimensions[0]!.levels[0]!;
		expect(level.points.length).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS + 1);
	});

	it('does not expand a hand-built level', () => {
		expect(chartExLevelValues({ pointCount: HUGE, points: [] }).length).toBeLessThanOrEqual(
			MAX_CHART_BLANK_POINTS,
		);
	});
});
