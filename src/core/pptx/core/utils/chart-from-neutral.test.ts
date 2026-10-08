/**
 * The pptx chart adapter over real chart XML: the part is read by the neutral parser
 * (`parseChartSpace`) and by fast-xml-parser (the styles), and the series and axes are built as
 * `getChartDataForGraphicFrame` builds them. The adapter was proven deeply equal to the previous
 * object-tree parser on every committed chart (81 frames) before that parser was deleted; the
 * corpus stays covered by `__tests__/integration/chart-neutral-parity.test.ts`. The cases below
 * are where the adapter deliberately reads differently (ECMA-376 `CT_Boolean` defaults, repeated
 * chart groups), plus the colour pipeline that used to regress for plain line charts.
 */
import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../../../chart/index';
import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { PptxChartData, XmlObject } from '../types';
import { parseChartAxes } from './chart-axis-parser';
import {
	axesFromNeutral,
	chartTypeFromNeutral,
	groupChartType,
	neutralCategories,
	pairGroupNodes,
	rawChildren,
	seriesFromNeutral,
	type ChartTreeReader,
} from './chart-from-neutral';

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
	extractPointValues: () => [],
};

const part = (plotArea: string) =>
	`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>${plotArea}</c:plotArea></c:chart></c:chartSpace>`;
const values =
	'<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>';
const outline = (color: string) =>
	`<c:spPr><a:ln><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr>`;
const fill = (color: string) =>
	`<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
const ser = (inner: string) => `<c:ser><c:idx val="0"/><c:order val="0"/>${inner}${values}</c:ser>`;

/** The chart's series and axes, built the way the runtime builds them. */
function build(xml: string): Pick<PptxChartData, 'chartType' | 'series' | 'axes'> {
	const { chartSpace } = parseChartSpace(xml);
	const plotArea = lookup.getChildByLocalName(
		lookup.getChildByLocalName(
			lookup.getChildByLocalName(parser.parse(xml), 'chartSpace'),
			'chart',
		),
		'plotArea',
	)!;
	const chartType = chartTypeFromNeutral(chartSpace);
	const { categories } = neutralCategories(chartSpace.plotArea.groups);
	const series = pairGroupNodes(chartSpace.plotArea.groups, plotArea, reader.localName).flatMap(
		({ group, node }) => {
			const type = groupChartType(group);
			return group.series.map((entry, index) =>
				seriesFromNeutral(
					{
						series: entry,
						node: rawChildren(node, 'ser', reader.localName)[index] ?? {},
						index,
						categories,
						...(chartType === 'combo' && type ? { seriesChartType: type } : {}),
						...(type ? { containerChartType: type } : {}),
					},
					reader,
				),
			);
		},
	);
	const axes = axesFromNeutral(
		chartSpace.plotArea.axes,
		parseChartAxes(plotArea, lookup, reader, reader.localName),
	);
	return { chartType, series, axes };
}

describe('chart adapter: series colour pipeline', () => {
	it.each([
		['lineChart', 'ED7D31'],
		['scatterChart', '4472C4'],
		['radarChart', '70AD47'],
		['stockChart', 'FFC000'],
	])('reads a plain %s series colour from a:ln/a:solidFill', (element, color) => {
		const { series } = build(part(`<c:${element}>${ser(outline(color))}</c:${element}>`));
		expect(series[0]?.color).toBe(`#${color}`);
		// Not a combo: no round-trip tag.
		expect(series[0]?.seriesChartType).toBeUndefined();
	});

	it('combo chart: bar keeps its direct fill, line its outline, both tagged', () => {
		const { chartType, series } = build(
			part(
				`<c:barChart>${ser(fill('4472C4'))}</c:barChart><c:lineChart>${ser(outline('ED7D31'))}</c:lineChart>`,
			),
		);
		expect(chartType).toBe('combo');
		expect(series.map((entry) => [entry.color, entry.seriesChartType])).toEqual([
			['#4472C4', 'bar'],
			['#ED7D31', 'line'],
		]);
	});
});

describe('chart adapter: reads where the object-tree parser did not', () => {
	it('reads every group of a repeated element (two c:lineChart groups on two axes)', () => {
		// fast-xml-parser groups same-named siblings into one array, which the object-tree reader
		// could not walk: the series of both groups were lost.
		const { series } = build(
			part(
				`<c:lineChart>${ser(outline('111111'))}</c:lineChart><c:lineChart>${ser(outline('222222'))}</c:lineChart>`,
			),
		);
		expect(series.map((entry) => entry.color)).toEqual(['#111111', '#222222']);
	});

	it('reads bare CT_Boolean switches as true (the schema default)', () => {
		const { series, axes } = build(
			part(
				`<c:lineChart>${ser('<c:marker/><c:smooth/>')}<c:axId val="1"/><c:axId val="2"/></c:lineChart>` +
					'<c:catAx><c:axId val="1"/><c:scaling/><c:delete/><c:crossAx val="2"/></c:catAx>' +
					'<c:valAx><c:axId val="2"/><c:scaling/><c:crossAx val="1"/></c:valAx>',
			),
		);
		expect(series[0]?.smooth).toBe(true);
		// A marker element without c:symbol is PowerPoint's automatic marker.
		expect(series[0]?.marker).toEqual({ symbol: 'auto' });
		expect(axes?.map((axis) => [axis.axisType, axis.deleted])).toEqual([
			['catAx', true],
			['valAx', undefined],
		]);
	});

	it('keeps the pptx axis order: grouped by element, each in document order', () => {
		const { axes } = build(
			part(
				'<c:valAx><c:axId val="1"/><c:scaling/></c:valAx><c:catAx><c:axId val="2"/><c:scaling/></c:catAx><c:valAx><c:axId val="3"/><c:scaling/></c:valAx>',
			),
		);
		expect(axes?.map((axis) => axis.axisId)).toEqual([1, 3, 2]);
	});
});

describe('neutralCategories', () => {
	const categories = (plotArea: string) =>
		neutralCategories(parseChartSpace(part(plotArea)).chartSpace.plotArea.groups);

	it('expands a sparse cache to c:ptCount with blank slots', () => {
		expect(
			categories(
				'<c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v> a </c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat></c:ser></c:barChart>',
			),
		).toEqual({ categories: ['a', '', 'c', ''] });
	});

	it('forward-fills the outer levels of a multi-level category axis', () => {
		expect(
			categories(
				'<c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="3"/><c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Apr</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="2"><c:v>Q2</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart>',
			),
		).toEqual({
			categories: ['Jan', 'Feb', 'Apr'],
			categoryLevels: [
				['Jan', 'Feb', 'Apr'],
				['Q1', 'Q1', 'Q2'],
			],
		});
	});

	it('falls back to number categories, then to the x values of the first series', () => {
		expect(
			categories(
				'<c:scatterChart><c:ser><c:xVal><c:strRef><c:strCache><c:pt idx="0"><c:v>x</c:v></c:pt></c:strCache></c:strRef></c:xVal></c:ser></c:scatterChart>',
			),
		).toEqual({ categories: ['x'] });
		expect(
			categories(
				'<c:barChart><c:ser><c:cat><c:numRef><c:numCache><c:pt idx="0"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:cat></c:ser></c:barChart>',
			),
		).toEqual({ categories: ['7'] });
	});
});
