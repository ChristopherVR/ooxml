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
import { applyChartStylePreset } from './edit/style';
import { setChartSeriesValues } from './edit/series';
import { addChartCategory } from './edit/series-rows';
import {
	CHART_POINTS_LIMITED,
	MAX_CACHE_BLANK_POINTS,
	MAX_CHART_BLANK_POINTS,
	chartPointBudget,
} from './point-limit';
import { chartExSummary } from './render/summary-from-chartex';
import { renderChartSummarySvg } from './render/render-space';
import { chartSummaryFromSpace } from './render/summary-from-space';
import { chartSpaceWorkbookWrites } from './workbook/space-writes';
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
		expect(counted.length).toBeLessThanOrEqual(MAX_CACHE_BLANK_POINTS);
		const indexed = chartCacheValues({ type: 'string', points: [{ index: HUGE, value: 'x' }] });
		expect(indexed.length).toBeLessThanOrEqual(MAX_CACHE_BLANK_POINTS + 1);
	});

	it('shows a large sparse range in full and keeps it editable', () => {
		const xml = barChart(
			series(
				0,
				'<c:ptCount val="199999"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="4998"><c:v>7</c:v></c:pt>',
			),
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(issues).toEqual([]);
		const values = chartCacheValues(chartSpace.plotArea.groups[0]!.series[0]!.values!.cache);
		expect(values).toHaveLength(199_999);
		expect(values[4998]).toBe(7);
		expect(chartIssuesLostOnWrite(issues)).toEqual([]);
	});

	it('keeps a huge declared count and shows only what fits', () => {
		const xml = barChart(series(0, `<c:ptCount val="${HUGE}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`));
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(codes(issues)).toEqual([CHART_POINTS_LIMITED]);
		expect(chartIssuesLostOnWrite(issues)).toEqual([]);
		expect(firstValues(xml).cache!.pointCount).toBe(HUGE);
		expect(chartSummaryFromSpace(chartSpace)!.series[0]!.values).toEqual([1]);
		expect(writeChartSpace(chartSpace)).toContain(`<c:ptCount val="${HUGE}"/>`);
	});

	it('keeps a point past the view in the model and in the saved part', () => {
		const xml = barChart(
			series(
				0,
				`<c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="${HUGE}"><c:v>2</c:v></c:pt>`,
			),
		);
		const values = firstValues(xml);
		expect(codes(parseChartSpace(xml).issues)).toEqual([CHART_POINTS_LIMITED]);
		expect(chartCacheValues(values.cache)).toEqual([1, '']);
		expect(values.cache!.points).toHaveLength(2);
		expect(writeChartSpace(parseChartSpace(xml).chartSpace)).toContain(`<c:pt idx="${HUGE}">`);
	});

	it('shares one budget among the caches of a part', () => {
		const count = 2000;
		const xml = barChart(
			Array.from({ length: count }, (_, index) =>
				series(
					index,
					`<c:ptCount val="${MAX_CACHE_BLANK_POINTS}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`,
				),
			).join(''),
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(codes(issues)).toEqual([CHART_POINTS_LIMITED]);
		const lengths = chartSpace.plotArea.groups[0]!.series.map(
			(entry) => chartCacheValues(entry.values!.cache).length,
		);
		expect(lengths.reduce((sum, length) => sum + length, 0)).toBeLessThanOrEqual(
			MAX_CHART_BLANK_POINTS + count,
		);
		expect(lengths.at(-1)).toBe(1);
		expect(chartSpaceDataSheet(chartSpace).categories).toHaveLength(1);
	});

	it('shares one budget among the parts of a package', () => {
		const pointBudget = chartPointBudget();
		const part = barChart(
			series(0, `<c:ptCount val="${MAX_CACHE_BLANK_POINTS}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`),
		);
		const lengths = Array.from({ length: 10 }, () => {
			const { chartSpace } = parseChartSpace(part, { pointBudget });
			return chartCacheValues(chartSpace.plotArea.groups[0]!.series[0]!.values!.cache).length;
		});
		expect(lengths.filter((length) => length === 1)).toHaveLength(6);
		expect(pointBudget.cuts).toBe(6);
	});

	it('does not depend on how the chart namespace is prefixed', () => {
		const xml = barChart(series(0, `<c:ptCount val="${HUGE}"/><c:pt idx="0"><c:v>1</c:v></c:pt>`))
			.replaceAll('<c:', '<é:')
			.replaceAll('</c:', '</é:')
			.replace('xmlns:c=', 'xmlns:é=');
		expect(firstValues(xml).cache!.pointCount).toBe(HUGE);
		expect(chartCacheValues(firstValues(xml).cache)).toEqual([1]);
	});

	it('keeps the declared count of a multi-level category cache and bounds its levels', () => {
		const levels = '<c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt></c:lvl>'.repeat(50);
		const categories = `<c:cat><c:multiLvlStrRef><c:f>S!$A$1</c:f><c:multiLvlStrCache><c:ptCount val="${HUGE}"/>${levels}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
		const { chartSpace, issues } = parseChartSpace(
			barChart(series(0, '<c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>', categories)),
		);
		const source = chartSpace.plotArea.groups[0]!.series[0]!.categories!;
		expect(codes(issues)).toEqual([CHART_POINTS_LIMITED]);
		expect(source.levelPointCount).toBe(HUGE);
		expect(source.levels!.every((level) => level.viewLength === 1)).toBe(true);
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

describe('saving a chart whose view is cut', () => {
	const cutPart = barChart(
		series(
			0,
			`<c:ptCount val="${HUGE}"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="${HUGE - 1}"><c:v>9</c:v></c:pt>`,
		),
	);

	it('keeps the source bytes when nothing changed, under any budget', () => {
		const { chartSpace } = parseChartSpace(cutPart, { pointBudget: chartPointBudget(0) });
		expect(writeChartSpaceOverSource(cutPart, chartSpace)).toBe(cutPart);
	});

	it('keeps the count and every point through a formatting edit', () => {
		const { chartSpace } = parseChartSpace(cutPart);
		const saved = writeChartSpaceOverSource(cutPart, { ...chartSpace, roundedCorners: true });
		expect(saved).toContain(`<c:ptCount val="${HUGE}"/>`);
		expect(saved).toContain(`<c:pt idx="${HUGE - 1}">`);
		expect(saved.match(/<c:pt /g)).toHaveLength(2);
	});

	it('applies a value edit to the view and keeps what lies past it', () => {
		const { chartSpace } = parseChartSpace(cutPart);
		const view = chartSpaceDataSheet(chartSpace).series[0]!.values;
		const edited = setChartSeriesValues(
			chartSpace,
			0,
			view.map(() => 5),
		);
		const cache = edited.plotArea.groups[0]!.series[0]!.values!.cache!;
		expect(cache.pointCount).toBe(HUGE);
		expect(cache.points).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ index: 0, value: '5' }),
				expect.objectContaining({ index: HUGE - 1, value: '9' }),
			]),
		);
	});

	it('adds a category to a cut view inside the declared range', () => {
		const { chartSpace } = parseChartSpace(cutPart);
		const cache = addChartCategory(chartSpace, 'New').plotArea.groups[0]!.series[0]!.values!.cache!;
		expect(cache.pointCount).toBe(HUGE);
		expect(cache.points.at(-1)).toMatchObject({ index: HUGE - 1, value: '9' });
	});

	it('keeps the ChartEx source bytes, and its count through an edit', () => {
		const xml = chartEx(`<cx:lvl ptCount="${HUGE}"><cx:pt idx="0">1</cx:pt></cx:lvl>`);
		const { chartExSpace } = parseChartExSpace(xml);
		expect(writeChartExSpaceOverSource(xml, chartExSpace)).toBe(xml);
		const saved = writeChartExSpaceOverSource(xml, {
			...chartExSpace,
			title: { tx: { value: 'New' } },
		});
		expect(saved).toContain(`ptCount="${HUGE}"`);
	});

	it('adds format overrides only for points that exist', () => {
		const pie = barChart(series(0, '<c:ptCount val="100000"/><c:pt idx="7"><c:v>1</c:v></c:pt>'))
			.replace('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>', '<c:pieChart>')
			.replace('<c:axId val="1"/><c:axId val="2"/></c:barChart>', '</c:pieChart>');
		const styled = applyChartStylePreset(parseChartSpace(pie).chartSpace, 'colorful');
		expect(styled.plotArea.groups[0]!.series[0]!.dataPoints.map((point) => point.index)).toEqual([
			7,
		]);
	});
});

describe('embedded workbook writes of a sparse chart', () => {
	const part = barChart(
		series(
			0,
			'<c:ptCount val="8"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>',
		),
	);
	const space = () => parseChartSpace(part).chartSpace;

	it('leaves the blank cells of a range alone', () => {
		const [write] = chartSpaceWorkbookWrites(space());
		expect(write).toMatchObject({ formula: 'S!$B$2:$B$9', cellCount: 8 });
		expect(chartSpaceWorkbookWrites(space(), space())).toEqual([]);
		expect(write!.cells).toEqual([
			{ index: 0, value: '1' },
			{ index: 2, value: '3' },
		]);
	});

	it('empties the cell of a point the edit removed', () => {
		const edited = setChartSeriesValues(space(), 0, [1, null, null, null, null, null, null, null]);
		expect(chartSpaceWorkbookWrites(edited, space())[0]!.cells).toEqual([{ index: 2, value: '' }]);
	});

	it('names every cell of the range when a cut view reaches only the first', () => {
		const { chartSpace } = parseChartSpace(part, { pointBudget: chartPointBudget(0) });
		const [write] = chartSpaceWorkbookWrites(chartSpace);
		expect(write!.cellCount).toBe(8);
		expect(write!.cells.map((cell) => cell.index)).toEqual([0, 2]);
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

	it('keeps a huge declared count and shows only what fits', () => {
		const xml = chartEx(`<cx:lvl ptCount="${HUGE}"><cx:pt idx="0">1</cx:pt></cx:lvl>`);
		const result = chartExSummary(xml)!;
		expect(result.pointsLimited).toBe(true);
		expect(result.summary!.series[0]!.values).toEqual([1]);
		const { chartExSpace, issues } = parseChartExSpace(xml);
		expect(codes(issues)).toEqual([CHART_POINTS_LIMITED]);
		expect(chartExSpace.chartData.data[0]!.dimensions[0]!.levels[0]!.pointCount).toBe(HUGE);
	});

	it('keeps a point past the view', () => {
		const xml = chartEx(
			`<cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt><cx:pt idx="${HUGE}">2</cx:pt></cx:lvl>`,
		);
		const result = chartExSummary(xml)!;
		expect(result.pointsLimited).toBe(true);
		expect(result.summary!.series[0]!.values).toEqual([1, null]);
		expect(writeChartExSpace(parseChartExSpace(xml).chartExSpace)).toContain(`idx="${HUGE}"`);
	});

	it('gives a new series zeros only where the level has points', () => {
		const { chartExSpace } = parseChartExSpace(
			chartEx(`<cx:lvl ptCount="${HUGE}"><cx:pt idx="0">1</cx:pt></cx:lvl>`),
		);
		expect(chartExSpaceDataSheet(chartExSpace).categories).toHaveLength(1);
		const added = addChartExSeries(chartExSpace, 'Copy');
		const level = added.chartData.data.at(-1)!.dimensions[0]!.levels[0]!;
		expect(level.points).toEqual([{ index: 0, value: '0' }]);
		expect(level.pointCount).toBe(HUGE);
	});

	it('does not expand a hand-built level', () => {
		expect(chartExLevelValues({ pointCount: HUGE, points: [] }).length).toBeLessThanOrEqual(
			MAX_CACHE_BLANK_POINTS,
		);
	});
});
