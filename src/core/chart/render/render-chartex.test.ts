import { describe, expect, it } from 'vitest';
import { chartExSummary } from './summary-from-chartex';
import { chartSummaryView } from './chart-view';
import { chartThemePalette, renderChartSummarySvg } from './render-space';

const CX = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const points = (values: (string | number)[]) =>
	`<cx:lvl ptCount="${values.length}">${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`;

/** A waterfall as Word or PowerPoint write it: formulas plus cached points. */
const waterfallXml = (options = '') => `<cx:chartSpace ${CX} ${A}><cx:chartData><cx:data id="0">
<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$7</cx:f>${points(['Start', 'Sales', 'Returns', 'Costs', 'Other', 'Total'])}</cx:strDim>
<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$7</cx:f>${points([100, 60, -20, -45, 15, 110])}</cx:numDim></cx:data></cx:chartData>
<cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>Cash flow</cx:v></cx:txData></cx:tx></cx:title>
<cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Amount</cx:v></cx:txData></cx:tx><cx:dataId val="0"/>
<cx:layoutPr>${options || '<cx:visibility connectorLines="1"/><cx:subtotals><cx:idx val="0"/><cx:idx val="5"/></cx:subtotals>'}</cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea>
<cx:legend pos="t" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>`;

const histogramXml = (binning: string) => `<cx:chartSpace ${CX} ${A}><cx:chartData><cx:data id="0">
<cx:numDim type="val"><cx:f>Sheet1!$H$2:$H$21</cx:f>${points([3, 5, 7, 8, 8, 9, 12, 13, 13, 14, 15, 15, 15, 17, 18, 21, 22, 24, 27, 29])}</cx:numDim></cx:data></cx:chartData>
<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:tx><cx:txData><cx:v>Score</cx:v></cx:txData></cx:tx><cx:dataId val="0"/>
<cx:layoutPr>${binning}</cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;

describe('chartExSummary', () => {
	it('summarises a waterfall: cached points, references, subtotals and title', () => {
		const result = chartExSummary(waterfallXml());
		expect(result?.layoutId).toBe('waterfall');
		const summary = result!.summary!;
		expect(summary).toMatchObject({
			chartType: 'waterfall',
			title: 'Cash flow',
			showLegend: true,
			legendPosition: 't',
			waterfall: { subtotalIndices: [0, 5] },
		});
		expect(summary.series[0]).toMatchObject({
			name: 'Amount',
			nameRef: 'Sheet1!$B$1',
			categoriesRef: 'Sheet1!$A$2:$A$7',
			valuesRef: 'Sheet1!$B$2:$B$7',
			values: [100, 60, -20, -45, 15, 110],
		});
	});

	it('reads connector visibility and leaves subtotals unset when absent', () => {
		const hidden = chartExSummary(waterfallXml('<cx:visibility connectorLines="0"/>'))!;
		expect(hidden.summary?.waterfall).toEqual({ connectorLines: false });
	});

	it('reads histogram binning as attributes or child values', () => {
		const attrs = chartExSummary(
			histogramXml('<cx:binning intervalClosed="r"><cx:binSize val="10"/></cx:binning>'),
		)!;
		expect(attrs.summary).toMatchObject({
			chartType: 'histogram',
			histogram: { intervalClosed: 'r', binSize: 10 },
		});
		const counted = chartExSummary(
			histogramXml('<cx:binning intervalClosed="l"><cx:binCount val="4"/></cx:binning>'),
		)!;
		expect(counted.summary?.histogram).toMatchObject({ binCount: 4 });
	});

	it('names an extended family it does not summarise, and ignores classic parts', () => {
		const map = waterfallXml().replace('layoutId="waterfall"', 'layoutId="regionMap"');
		expect(chartExSummary(map)).toEqual({ layoutId: 'regionMap' });
		expect(chartExSummary('<c:chartSpace xmlns:c="urn:x"/>')).toBeUndefined();
	});
});

describe('waterfall painting', () => {
	const summary = chartExSummary(waterfallXml())!.summary!;

	it('floats each bar from the running total and colours it by role', () => {
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		expect(view.waterfall?.steps.map((s) => [s.startValue, s.endValue, s.isSubtotal])).toEqual([
			[0, 100, true],
			[100, 160, false],
			[160, 140, false],
			[140, 95, false],
			[95, 110, false],
			[0, 110, true],
		]);
		expect(view.valueAxis?.max).toBeGreaterThanOrEqual(160);
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		const bars = svg.match(/<g data-chart-series="0" data-chart-point="\d"><rect [^>]*>/g)!;
		expect(bars).toHaveLength(6);
		const fills = bars.map((b) => /fill="(#[0-9A-Fa-f]+)"/.exec(b)![1]);
		// Totals share one colour; increases and decreases differ from it and from each other.
		expect(fills[0]).toBe(fills[5]);
		expect(new Set(fills).size).toBe(3);
		expect(fills[1]).toBe(fills[4]);
		expect(fills[2]).toBe(fills[3]);
		// Five dashed connectors between six bars; the fixed three-entry legend.
		expect(svg.match(/stroke-dasharray/g)).toHaveLength(5);
		for (const label of ['Increase', 'Decrease', 'Total', 'Cash flow', 'Returns'])
			expect(svg).toContain(`>${label}<`);
	});

	it('omits connectors when the chart hides them and treats only the last point as a total by default', () => {
		const noConnectors = chartExSummary(
			waterfallXml('<cx:visibility connectorLines="0"/>'),
		)!.summary!;
		const svg = renderChartSummarySvg(noConnectors, { width: 480, height: 300 });
		expect(svg).not.toContain('stroke-dasharray');
		const view = chartSummaryView(noConnectors, chartThemePalette(undefined));
		expect(view.waterfall?.steps.filter((s) => s.isSubtotal).map((s) => s.sourceIndex)).toEqual([
			5,
		]);
	});
});

describe('histogram painting', () => {
	it('bins the observations and draws one gapless count bar per bin', () => {
		const summary = chartExSummary(histogramXml('<cx:binning intervalClosed="r"/>'))!.summary!;
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		const total = view.series[0]!.values.reduce<number>((sum, v) => sum + (v ?? 0), 0);
		expect(total).toBe(20);
		expect(view.categories.length).toBeGreaterThan(1);
		expect(view.categories[0]).toMatch(/^[[(]/);
		expect(view.barGapWidth).toBe(0);
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		expect(svg.match(/<g data-chart-series="0" data-chart-point/g)).toHaveLength(
			view.categories.length,
		);
		expect(svg).not.toContain('not drawn');
	});

	it('honours an explicit bin size', () => {
		const summary = chartExSummary(
			histogramXml('<cx:binning intervalClosed="l"><cx:binSize val="10"/></cx:binning>'),
		)!.summary!;
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		expect(view.series[0]!.values).toEqual([6, 9, 5]);
	});
});
