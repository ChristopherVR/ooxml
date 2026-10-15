import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { chartSummaryView } from './chart-view';
import { chartSpaceKind, chartThemePalette, renderChartSpaceSvg } from './render-space';
import { chartExSummary } from './summary-from-chartex';
import { chartSummaryFromSpace } from './summary-from-space';
import { renderChartSummarySvg } from './render-space';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const CX = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';

const num = (values: number[]) =>
	`<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}`;
const name = (text: string) =>
	`<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${text}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
const cats = ['Mon', 'Tue', 'Wed', 'Thu'];
const catXml = `<c:cat><c:strRef><c:f>S!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;

const axes = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`;

const stockSeries = (idx: number, label: string, values: number[]) =>
	`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${name(label)}${catXml}<c:val><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache>${num(values)}</c:numCache></c:numRef></c:val></c:ser>`;

const stockXml = (series: string) => `<c:chartSpace ${C}><c:chart><c:plotArea><c:layout/>
<c:stockChart>${series}<c:hiLowLines/><c:axId val="1"/><c:axId val="2"/></c:stockChart>${axes}</c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`;

const hlc = [
	stockSeries(0, 'High', [14, 15, 15, 14]),
	stockSeries(1, 'Low', [9, 10, 10, 9]),
	stockSeries(2, 'Close', [12, 13, 11, 13]),
].join('');
const ohlc =
	stockSeries(0, 'Open', [10, 11, 12, 11]) +
	hlc.replace('<c:idx val="2"/><c:order val="2"/>', '<c:idx val="3"/><c:order val="3"/>');

const bubbleSeries = (idx: number, label: string, xs: number[], ys: number[], sizes: number[]) =>
	`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${name(label)}<c:xVal><c:numRef><c:f>S!$A$2:$A$5</c:f><c:numCache>${num(xs)}</c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache>${num(ys)}</c:numCache></c:numRef></c:yVal><c:bubbleSize><c:numRef><c:f>S!$C$2:$C$5</c:f><c:numCache>${num(sizes)}</c:numCache></c:numRef></c:bubbleSize></c:ser>`;

const bubbleXml = (
	options = '',
	sizes = [5, 10, 3, 8],
) => `<c:chartSpace ${C}><c:chart><c:plotArea><c:layout/>
<c:bubbleChart>${bubbleSeries(0, 'A', [1, 2, 3, 4], [3, 5, 4, 8], sizes)}${bubbleSeries(1, 'B', [1.5, 2.5, 3.5, 4.5], [6, 2, 9, 1], [2, 4, 9, 1])}${options}<c:axId val="1"/><c:axId val="2"/></c:bubbleChart>
<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
</c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`;

const render = (xml: string) =>
	renderChartSpaceSvg(parseChartSpace(xml).chartSpace, { width: 480, height: 300 });
const radii = (svg: string) =>
	[...svg.matchAll(/<circle [^>]*\br="([\d.]+)"[^>]*fill-opacity="0.6"/g)].map((m) => Number(m[1]));

describe('bubble charts', () => {
	it('reads sizes and display options into the summary', () => {
		const summary = chartSummaryFromSpace(
			parseChartSpace(
				bubbleXml(
					'<c:bubbleScale val="150"/><c:showNegBubbles val="1"/><c:sizeRepresents val="w"/>',
				),
			).chartSpace,
		);
		expect(summary.chartType).toBe('bubble');
		expect(summary.series[0]?.bubbleSizes).toEqual([5, 10, 3, 8]);
		expect(summary.series[0]?.bubbleSizesRef).toBe('S!$C$2:$C$5');
		expect(summary.bubble).toEqual({
			sizeRepresents: 'w',
			bubbleScale: 150,
			showNegativeBubbles: true,
		});
		const defaults = chartSummaryFromSpace(parseChartSpace(bubbleXml()).chartSpace);
		expect(defaults.bubble).toEqual({ sizeRepresents: 'area', showNegativeBubbles: false });
	});

	it('draws one translucent circle per point, larger for bigger sizes, with a legend', () => {
		const { chartSpace } = parseChartSpace(bubbleXml());
		expect(chartSpaceKind(chartSpace)).toEqual({ type: 'bubble', drawn: true });
		const svg = render(bubbleXml());
		const r = radii(svg);
		expect(r).toHaveLength(8);
		// Series A sizes 5, 10, 3, 8: radius follows size; the shared maximum is series B's 9... A's 10.
		expect(r[1]).toBeGreaterThan(r[3]!);
		expect(r[3]).toBeGreaterThan(r[0]!);
		expect(r[0]).toBeGreaterThan(r[2]!);
		expect(svg).toContain('>A<');
		expect(svg).toContain('>B<');
		expect(svg).not.toContain('not drawn');
		expect(svg).not.toMatch(/NaN|undefined/);
	});

	it('hides negative bubbles unless the chart shows them', () => {
		expect(radii(render(bubbleXml('', [5, -10, 3, 8])))).toHaveLength(7);
		expect(radii(render(bubbleXml('<c:showNegBubbles val="1"/>', [5, -10, 3, 8])))).toHaveLength(8);
	});

	it('sizes by width rather than area when asked', () => {
		const area = radii(render(bubbleXml('<c:sizeRepresents val="area"/>')));
		const width = radii(render(bubbleXml('<c:sizeRepresents val="w"/>')));
		// The smallest bubble (size 1 of 10) is relatively bigger under area scaling.
		expect(Math.min(...area)).toBeGreaterThan(Math.min(...width));
	});

	it('prefers live bubble sizes when the host resolves the reference', () => {
		const summary = chartSummaryFromSpace(parseChartSpace(bubbleXml()).chartSpace);
		const view = chartSummaryView(summary, chartThemePalette(undefined), {
			evaluateRef: (ref) => (ref === 'S!$C$2:$C$5' ? [1, 2, 3, 4] : []),
		});
		expect(view.series[0]?.sizes).toEqual([1, 2, 3, 4]);
	});
});

describe('stock charts', () => {
	it('draws a wick and a close tick per category for high-low-close', () => {
		const { chartSpace } = parseChartSpace(stockXml(hlc));
		expect(chartSpaceKind(chartSpace)).toEqual({ type: 'stock', drawn: true });
		const svg = render(stockXml(hlc));
		expect(svg.match(/<g data-chart-series="2" data-chart-point="\d"/g)).toHaveLength(4);
		// Each candle group: one wick and one tick, no body rectangle.
		const groups = svg.match(/<g data-chart-series="2"[\s\S]*?<\/g>/g)!;
		for (const group of groups) {
			expect(group.match(/<line /g)).toHaveLength(2);
			expect(group).not.toContain('<rect');
		}
		expect(svg).toContain('>High<');
		expect(svg).toContain('>Close<');
	});

	it('draws up and down bodies from open to close for open-high-low-close', () => {
		const svg = render(stockXml(ohlc));
		const groups = svg.match(/<g data-chart-series="3"[\s\S]*?<\/g>/g)!;
		expect(groups).toHaveLength(4);
		const fills = groups.map((g) => /<rect [^>]*fill="(#[0-9A-F]+)"/.exec(g)![1]);
		// Closes 12, 13, 11, 13 against opens 10, 11, 12, 11: up, up, down, up.
		expect(fills).toEqual(['#FFFFFF', '#FFFFFF', '#404040', '#FFFFFF']);
		for (const group of groups) expect(group.match(/<line /g)).toHaveLength(1);
	});

	it('scales the value axis over the highs and lows', () => {
		const view = chartSummaryView(
			chartSummaryFromSpace(parseChartSpace(stockXml(ohlc)).chartSpace),
			chartThemePalette(undefined),
		);
		expect(view.supported).toBe(true);
		expect(view.valueAxis!.max).toBeGreaterThanOrEqual(15);
	});
});

const points = (values: (string | number)[]) =>
	`<cx:lvl ptCount="${values.length}">${values.map((v, i) => (v === '' ? '' : `<cx:pt idx="${i}">${v}</cx:pt>`)).join('')}</cx:lvl>`;

const hierarchyXml = (
	layoutId: string,
	layoutPr = '',
) => `<cx:chartSpace ${CX}><cx:chartData><cx:data id="0">
<cx:strDim type="cat"><cx:f>S!$A$2:$B$6</cx:f>${points(['Apple', 'Pear', 'Leek', 'Kale', 'Almond'])}${points(['Fruit', '', 'Veg', '', 'Nuts'])}</cx:strDim>
<cx:numDim type="size"><cx:f>S!$C$2:$C$6</cx:f>${points([30, 20, 10, 5, 8])}</cx:numDim></cx:data></cx:chartData>
<cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>Produce</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layoutId}"><cx:tx><cx:txData><cx:v>Weight</cx:v></cx:txData></cx:tx><cx:dataId val="0"/>${layoutPr}</cx:series></cx:plotAreaRegion></cx:plotArea><cx:legend pos="b" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>`;

describe('treemap charts', () => {
	it('reads the category levels and the parent label layout', () => {
		const result = chartExSummary(
			hierarchyXml(
				'treemap',
				'<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>',
			),
		)!;
		expect(result.layoutId).toBe('treemap');
		expect(result.summary).toMatchObject({
			chartType: 'treemap',
			title: 'Produce',
			hierarchy: {
				levels: [
					['Apple', 'Pear', 'Leek', 'Kale', 'Almond'],
					['Fruit', '', 'Veg', '', 'Nuts'],
				],
				parentLabelLayout: 'overlapping',
			},
		});
		expect(result.summary!.series[0]!.values).toEqual([30, 20, 10, 5, 8]);
		expect(result.summary!.series[0]!.valuesRef).toBe('S!$C$2:$C$6');
	});

	it('draws a cell per leaf coloured by branch, parent labels and a branch legend', () => {
		const summary = chartExSummary(hierarchyXml('treemap'))!.summary!;
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		expect(view.hierarchy?.branches).toEqual(['Fruit', 'Veg', 'Nuts']);
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		expect(svg.match(/<g data-chart-series="0" data-chart-point="\d"/g)).toHaveLength(5);
		const fills = [
			...svg.matchAll(
				/<g data-chart-series="0" data-chart-point="(\d)"><rect [^>]*fill="(#[0-9A-F]+)"/g,
			),
		];
		const byPoint = new Map(fills.map((m) => [m[1], m[2]]));
		// Apple and Pear share the Fruit colour; Leek and Kale share Veg; the three differ.
		expect(byPoint.get('0')).toBe(byPoint.get('1'));
		expect(byPoint.get('2')).toBe(byPoint.get('3'));
		expect(new Set(byPoint.values()).size).toBe(3);
		for (const label of ['Fruit', 'Veg', 'Nuts', 'Produce']) expect(svg).toContain(`>${label}<`);
		expect(svg).not.toContain('not drawn');
	});

	it('draws a flat treemap with a cell per category', () => {
		const flat = hierarchyXml('treemap').replace(
			/<cx:lvl ptCount="5">(?:<cx:pt idx="[024]">(?:Fruit|Veg|Nuts)<\/cx:pt>)+<\/cx:lvl>/,
			'',
		);
		const summary = chartExSummary(flat)!.summary!;
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		expect(svg.match(/<g data-chart-series="0" data-chart-point="\d"/g)).toHaveLength(5);
	});
});

describe('sunburst charts', () => {
	it('draws a ring per level with wedges for every parent and leaf', () => {
		const summary = chartExSummary(hierarchyXml('sunburst'))!.summary!;
		expect(summary.chartType).toBe('sunburst');
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		// Three parent wedges plus five leaf wedges.
		expect(svg.match(/<path d="M [^"]*" fill="/g)).toHaveLength(8);
		expect(svg.match(/data-chart-point="\d"/g)).toHaveLength(5);
		for (const label of ['Fruit', 'Apple', 'Almond']) expect(svg).toContain(`>${label}</text>`);
		expect(svg).not.toMatch(/NaN|undefined/);
	});

	it('draws one ring per series for a flat sunburst', () => {
		const flat = hierarchyXml('sunburst').replace(
			/<cx:lvl ptCount="5">(?:<cx:pt idx="[024]">(?:Fruit|Veg|Nuts)<\/cx:pt>)+<\/cx:lvl>/,
			'',
		);
		const svg = renderChartSummarySvg(chartExSummary(flat)!.summary!, { width: 480, height: 300 });
		expect(svg.match(/<path d="M [^"]*" fill="/g)).toHaveLength(5);
	});
});
