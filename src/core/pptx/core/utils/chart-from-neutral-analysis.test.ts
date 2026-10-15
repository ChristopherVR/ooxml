/**
 * Trendlines and error bars built from the neutral chart model agree with the earlier object-tree
 * readers (`parseSeriesTrendlines`, `parseSeriesErrBars`, deprecated) field for field and in key
 * order, so the chart view models hashed by `ooxml-ui`'s chart parity test do not move. Where the
 * adapter reads differently it follows the schema, pinned below: `c:name` is read (the tree reader
 * never found its text), and bare `CT_Boolean` switches (`<c:dispRSqr/>`) read as true and a bare
 * `<c:trendlineLbl/>` as a label, which fast-xml-parser's empty strings hid.
 */
import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../../../chart/index';
import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { XmlObject } from '../types';
import { parseSeriesErrBars, parseSeriesTrendlines } from './chart-advanced-parser';
import { errBarsFromNeutral, trendlinesFromNeutral } from './chart-from-neutral-analysis';
import type { ChartTreeReader } from './chart-series-tree-fields';

const lookup = new PptxXmlLookupService();
const parser = new PptxRuntimeDependencyFactory().createParser();
const reader: ChartTreeReader = {
	xml: lookup,
	parseColor: (node) => {
		const value = (node?.['a:srgbClr'] as XmlObject | undefined)?.['@_val'];
		return typeof value === 'string' ? `#${value}` : undefined;
	},
	resolveTypeface: (raw) => raw,
	localName: (key) => key.slice(key.lastIndexOf(':') + 1),
};

/** The runtime's dense point reader (`extractChartPointValues`), for the deprecated reader. */
function pointValues(container: XmlObject | undefined): string[] {
	const ref =
		lookup.getChildByLocalName(container, 'numRef') ??
		lookup.getChildByLocalName(container, 'numLit');
	const cache = lookup.getChildByLocalName(ref, 'numCache') ?? ref;
	return [...lookup.getChildrenArrayByLocalName(cache, 'pt')]
		.sort(
			(a, b) => parseInt(String(a['@_idx'] || '0'), 10) - parseInt(String(b['@_idx'] || '0'), 10),
		)
		.map((pt) => String(lookup.getScalarChildByLocalName(pt, 'v') || '').trim())
		.filter((value) => value.length > 0);
}

const ln = (color: string, extra = '') =>
	`<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${extra}</a:ln></c:spPr>`;

const ANALYSIS =
	`<c:trendline>${ln('FF0000', '<a:prstDash val="dash"/>')}<c:trendlineType val="poly"/><c:order val="3"/>` +
	'<c:forward val="1.5"/><c:backward val="0.5"/><c:intercept val="2"/><c:dispRSqr val="0"/><c:dispEq val="1"/>' +
	'<c:trendlineLbl><c:layout><c:manualLayout><c:x val="0.1"/><c:y val="-0.05"/></c:manualLayout></c:layout><c:numFmt formatCode="#,##0.00" sourceLinked="0"/></c:trendlineLbl></c:trendline>' +
	'<c:trendline><c:trendlineType val="cubic"/></c:trendline>' +
	'<c:trendline><c:trendlineType val="movingAvg"/><c:period val="3"/></c:trendline>' +
	'<c:trendline><c:trendlineType val="exp"/><c:dispEq val="0"/></c:trendline>' +
	`<c:errBars><c:errDir val="x"/><c:errBarType val="minus"/><c:errValType val="cust"/><c:noEndCap val="0"/>` +
	'<c:plus><c:numLit><c:ptCount val="3"/><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v> </c:v></c:pt></c:numLit></c:plus>' +
	`<c:minus><c:numRef><c:f>S!$A$1:$A$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>0.5</c:v></c:pt><c:pt idx="1"><c:v>x</c:v></c:pt></c:numCache></c:numRef></c:minus>${ln('00FF00')}</c:errBars>` +
	'<c:errBars><c:errValType val="percentage"/><c:val val="5"/></c:errBars>' +
	'<c:errBars><c:errBarType val="plus"/></c:errBars>';

/** What the tree readers missed: a name, a bare boolean switch and a bare label. */
const SCHEMA_READS =
	'<c:trendline><c:name>Trend</c:name><c:trendlineType val="linear"/><c:dispRSqr/><c:trendlineLbl/></c:trendline>';

const part = (analysis: string) =>
	`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>${analysis}<c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`;

function sources(analysis = ANALYSIS) {
	const xml = part(analysis);
	const series = parseChartSpace(xml).chartSpace.plotArea.groups[0]!.series[0]!;
	const node = lookup.getChildByLocalName(
		lookup.getChildByLocalName(
			lookup.getChildByLocalName(
				lookup.getChildByLocalName(
					lookup.getChildByLocalName(parser.parse(xml), 'chartSpace'),
					'chart',
				),
				'plotArea',
			),
			'lineChart',
		),
		'ser',
	)!;
	return { series, node };
}

describe('trendlines and error bars from the neutral model', () => {
	it('match the object-tree readers, keys in the same order', () => {
		const { series, node } = sources();
		const trendlines = trendlinesFromNeutral(series.trendlines, node, reader);
		const errBars = errBarsFromNeutral(series.errorBars, node, reader);
		expect(JSON.stringify(trendlines)).toBe(
			JSON.stringify(parseSeriesTrendlines(node, lookup, reader)),
		);
		expect(JSON.stringify(errBars)).toBe(
			JSON.stringify(parseSeriesErrBars(node, lookup, pointValues, reader)),
		);
	});

	it('reads the structure from the model and the line style from the tree', () => {
		const { series, node } = sources();
		expect(trendlinesFromNeutral(series.trendlines, node, reader)).toStrictEqual([
			{
				trendlineType: 'polynomial',
				order: 3,
				forward: 1.5,
				backward: 0.5,
				intercept: 2,
				displayRSq: false,
				displayEq: true,
				color: '#FF0000',
				lineWidth: 2,
				lineDashStyle: 'dash',
				label: { layout: { x: 0.1, y: -0.05 }, numberFormatCode: '#,##0.00', sourceLinked: false },
			},
			{ trendlineType: 'movingAvg', period: 3 },
			{ trendlineType: 'exponential', displayEq: false },
		]);
		expect(errBarsFromNeutral(series.errorBars, node, reader)).toStrictEqual([
			{
				direction: 'x',
				barType: 'minus',
				valType: 'cust',
				noEndCap: false,
				customPlus: [1, 3],
				customMinus: [0.5],
				color: '#00FF00',
				width: 2,
			},
			{ direction: 'y', barType: 'both', valType: 'percentage', val: 5 },
		]);
	});

	it('reads what the tree readers missed, as the schema defines it', () => {
		const { series, node } = sources(SCHEMA_READS);
		expect(trendlinesFromNeutral(series.trendlines, node, reader)).toStrictEqual([
			{ trendlineType: 'linear', name: 'Trend', displayRSq: true, label: {} },
		]);
		expect(parseSeriesTrendlines(node, lookup, reader)).toStrictEqual([
			{ trendlineType: 'linear' },
		]);
	});
});
