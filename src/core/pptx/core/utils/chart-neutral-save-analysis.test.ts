import { describe, expect, it } from 'vitest';
import {
	parseChartSpace,
	writeChartSpaceOverSource,
	chartSeriesEntryAt,
	setChartType,
	setChartMarker,
	setChartPointOptions,
	setChartSeriesOptions,
} from '../../../chart';
import { COLUMN_CHART_XML } from '../../../chart/edit/test-support/column-chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { planNeutralChartAnalysis } from './chart-neutral-save-analysis';
import { same } from './chart-neutral-save-common';

const LINE = COLUMN_CHART_XML.replaceAll('barChart', 'lineChart')
	.replace('<c:barDir val="col"/>', '')
	.replace('<c:gapWidth val="219"/><c:overlap val="-27"/>', '')
	.replace('val="clustered"', 'val="standard"');
const BASE = (): PptxChartData =>
	({
		chartType: 'line',
		categories: ['North', 'South', 'East'],
		series: [
			{ name: 'Plan', idx: 0, values: [10, 20, 30] },
			{ name: 'Actual', idx: 1, values: [12, 18, 33] },
		],
	}) as PptxChartData;
const PAINT =
	'<c:spPr><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="60000"/></a:schemeClr></a:solidFill><a:round/><a:extLst><a:ext uri="line-private"/></a:extLst></a:ln><a:effectLst><a:softEdge rad="12700"/></a:effectLst></c:spPr>';
function run(xml: string, before: PptxChartData, after: PptxChartData) {
	const probe = structuredClone(after),
		edits: NeutralChartEdit[] = [];
	planNeutralChartAnalysis(before, probe, edits);
	expect(same(probe, before)).toBe(true);
	let space = parseChartSpace(xml).chartSpace;
	for (const edit of edits) space = edit(space);
	const saved = writeChartSpaceOverSource(xml, space);
	return { saved, space: parseChartSpace(saved).chartSpace };
}

describe('classic neutral analysis save claims', () => {
	it('adds every trendline parameter, custom errors, marker formatting and smooth without changing caches', () => {
		const before = BASE(),
			after = structuredClone(before);
		after.series[0]!.trendlines = [
			{
				trendlineType: 'polynomial',
				order: 3,
				name: 'Fit',
				forward: 2,
				backward: 1,
				intercept: 5,
				displayEq: true,
				displayRSq: true,
				color: '#123456',
				lineWidth: 2,
				lineDashStyle: 'dash',
				label: { layout: { x: 0.2, y: 0.3 }, numberFormatCode: '0.00', sourceLinked: false },
			},
		];
		after.series[0]!.errBars = [
			{
				direction: 'y',
				barType: 'both',
				valType: 'cust',
				customPlus: [1, 2, 3],
				customMinus: [2, 3, 4],
				noEndCap: true,
				color: '#654321',
				width: 1.5,
				dashStyle: 'dot',
			},
		];
		after.series[0]!.marker = {
			symbol: 'diamond',
			size: 9,
			spPr: { fillColor: '#FF0000', strokeWidth: 2 },
		};
		after.series[0]!.smooth = true;
		const { saved, space } = run(LINE, before, after),
			series = chartSeriesEntryAt(space, 0).series;
		expect(series.trendlines?.[0]).toMatchObject({
			trendlineType: 'poly',
			order: 3,
			name: 'Fit',
			forward: 2,
			backward: 1,
			intercept: 5,
			displayEquation: true,
			displayRSquared: true,
			label: {
				layout: { x: 0.2, y: 0.3 },
				numberFormat: { formatCode: '0.00', sourceLinked: false },
			},
		});
		expect(series.errorBars?.[0]).toMatchObject({
			valueType: 'cust',
			noEndCap: true,
			plus: {
				kind: 'numLit',
				cache: {
					points: [
						{ index: 0, value: '1' },
						{ index: 1, value: '2' },
						{ index: 2, value: '3' },
					],
				},
			},
		});
		expect(series.marker).toMatchObject({
			symbol: 'diamond',
			size: 9,
			spPr: { line: { widthEmu: 25400 } },
		});
		expect(series.smooth).toBe(true);
		expect(saved).toContain('Sheet1!$B$2:$B$4');
		expect(series.values).toEqual(
			chartSeriesEntryAt(parseChartSpace(LINE).chartSpace, 0).series.values,
		);
		expect(space.plotArea.groups[0]!.marker).toBeUndefined();
	});
	it('changes analysis geometry while preserving authored theme tokens, effects, labels and custom references', () => {
		const extra =
			`<c:trendline>${PAINT}<c:trendlineType val="poly"/><c:order val="2"/><c:trendlineLbl><c:numFmt formatCode="0.0"/><c:extLst><c:ext uri="label-private"/></c:extLst></c:trendlineLbl><c:extLst><c:ext uri="trend-private"/></c:extLst></c:trendline>` +
			`<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="cust"/><c:plus><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:plus><c:minus><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:minus>${PAINT}</c:errBars>`;
		const xml = LINE.replace('<c:cat>', extra + '<c:cat>'),
			before = BASE();
		before.series[0]!.trendlines = [
			{
				trendlineType: 'polynomial',
				order: 2,
				color: '#ED7D31',
				lineWidth: 1,
				label: { numberFormatCode: '0.0' },
			},
		];
		before.series[0]!.errBars = [
			{
				direction: 'y',
				barType: 'both',
				valType: 'cust',
				customPlus: [1, 2, 3],
				customMinus: [1],
				color: '#ED7D31',
				width: 1,
			},
		];
		const after = structuredClone(before);
		after.series[0]!.trendlines![0]!.lineWidth = 3;
		after.series[0]!.trendlines![0]!.order = 4;
		after.series[0]!.errBars![0]!.width = 2;
		const { saved, space } = run(xml, before, after);
		expect(saved).toContain('val="accent2"');
		expect(saved).toContain('line-private');
		expect(saved).toContain('label-private');
		expect(saved).toContain('trend-private');
		expect(saved).toContain('softEdge');
		expect(chartSeriesEntryAt(space, 0).series.errorBars?.[0]?.plus?.formula).toBe(
			'Sheet1!$D$2:$D$4',
		);
		const cleared = structuredClone(after);
		cleared.series[0]!.trendlines = [];
		cleared.series[0]!.errBars = [];
		const result = run(saved, after, cleared);
		expect(result.saved).not.toContain('<c:trendline>');
		expect(result.saved).not.toContain('<c:errBars>');
	});
	it('edits keyed point shape/gradient/marker values while retaining unrelated extension and theme data', () => {
		const xml = LINE.replace(
			'<c:cat>',
			`<c:dPt><c:idx val="1"/><c:marker><c:symbol val="circle"/><c:size val="6"/>${PAINT}</c:marker>${PAINT}<c:extLst><c:ext uri="point-private"/></c:extLst></c:dPt><c:cat>`,
		);
		const before = BASE();
		before.series[0]!.dataPoints = [
			{
				idx: 1,
				spPr: { strokeColor: '#ED7D31', strokeWidth: 1 },
				marker: { symbol: 'circle', size: 6, spPr: { strokeColor: '#ED7D31', strokeWidth: 1 } },
			},
		];
		const after = structuredClone(before);
		after.series[0]!.dataPoints![0]!.spPr!.strokeWidth = 3;
		after.series[0]!.dataPoints![0]!.marker!.size = 8;
		const { saved, space } = run(xml, before, after);
		expect(saved).toContain('point-private');
		expect(saved).toContain('line-private');
		expect(saved).toContain('val="accent2"');
		expect(chartSeriesEntryAt(space, 0).series.dataPoints[0]).toMatchObject({
			index: 1,
			spPr: { line: { widthEmu: 38100 } },
			marker: { size: 8 },
		});
	});
	it('refuses unsupported picture changes instead of claiming the point array', () => {
		const before = BASE(),
			after = structuredClone(before),
			edits: NeutralChartEdit[] = [];
		after.series[0]!.dataPoints = [{ idx: 1, picture: { applyToFront: true } } as never];
		const probe = structuredClone(after);
		planNeutralChartAnalysis(before, probe, edits);
		expect(same(probe, before)).toBe(false);
		expect(edits).toEqual([]);
	});
	it('keeps theme gradient tokens when only point gradient geometry changes', () => {
		const xml = COLUMN_CHART_XML.replace(
			'<c:cat>',
			'<c:dPt><c:idx val="1"/><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/><a:extLst><a:ext uri="gradient-private"/></a:extLst></a:gradFill></c:spPr></c:dPt><c:cat>',
		);
		const before = BASE();
		before.chartType = 'bar';
		before.series[0]!.dataPoints = [
			{
				idx: 1,
				gradientFill: {
					type: 'linear',
					angle: 0,
					scaled: false,
					stops: [
						{ position: 0, color: '#4472C4' },
						{ position: 100, color: '#ED7D31' },
					],
				},
			},
		];
		const after = structuredClone(before);
		after.series[0]!.dataPoints![0]!.gradientFill!.angle = 45;
		const { saved, space } = run(xml, before, after);
		expect(saved).toContain('gradient-private');
		expect(saved).toContain('val="accent1"');
		expect(saved).toContain('val="accent2"');
		expect(chartSeriesEntryAt(space, 0).series.dataPoints[0]?.spPr?.fill).toMatchObject({
			kind: 'gradient',
			angle: 45,
			scaled: false,
		});
	});
	it('reassigns an existing axis pair in place and rejects broken references', () => {
		const xml = LINE.replace(
			'</c:plotArea>',
			'<c:catAx><c:axId val="33"/><c:scaling/><c:axPos val="t"/><c:crossAx val="44"/></c:catAx><c:valAx><c:axId val="44"/><c:scaling/><c:axPos val="r"/><c:crossAx val="33"/></c:valAx></c:plotArea>',
		);
		const before = BASE(),
			after = structuredClone(before);
		after.series[1]!.axisId = 44;
		const { space } = run(xml, before, after);
		expect(space.plotArea.groups.map((group) => group.axisIds)).toEqual([
			[11, 22],
			[33, 44],
		]);
		expect(
			space.plotArea.groups.flatMap((group) => group.series).map((series) => series.index),
		).toEqual([0, 1]);
		const reset = structuredClone(after);
		delete reset.series[1]!.axisId;
		expect(
			run(writeChartSpaceOverSource(xml, space), after, reset).space.plotArea.groups[1]?.axisIds,
		).toEqual([11, 22]);
		expect(() =>
			run(xml.replace('<c:crossAx val="33"/>', '<c:crossAx val="999"/>'), before, after),
		).toThrow('coherent');
	});
	it('resets markers without changing sibling defaults and validates scalar point/series admission', () => {
		const space = parseChartSpace(LINE).chartSpace;
		const marked = setChartMarker(space, 0, { symbol: 'circle', size: 9 });
		expect(marked.plotArea.groups[0]!.marker).toBeUndefined();
		expect(
			chartSeriesEntryAt(setChartMarker(marked, 0, { size: null }), 0).series.marker?.size,
		).toBeUndefined();
		expect(
			chartSeriesEntryAt(setChartMarker(marked, 0, undefined), 0).series.marker,
		).toBeUndefined();
		expect(() => setChartSeriesOptions(space, 0, { explosion: 20 })).toThrow();
		expect(() => setChartPointOptions(space, 0, 1, { explosion: 20 })).toThrow();
		const pie = setChartType(space, { kind: 'pie' });
		const edited = setChartPointOptions(setChartSeriesOptions(pie, 0, { explosion: 15 }), 0, 1, {
			explosion: 25,
		});
		expect(chartSeriesEntryAt(edited, 0).series).toMatchObject({
			explosion: 15,
			dataPoints: [{ index: 1, explosion: 25 }],
		});
	});
});
