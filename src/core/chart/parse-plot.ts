// `c:chartSpace` parser: the plot area, its chart groups and its axes.
import { NS, children, elements, type XmlElement } from '../xml/index';
import type {
	ChartAxis,
	ChartAxisKind,
	ChartAxisScaling,
	ChartGroupKind,
	ChartPlotArea,
	ChartPlotGroup,
} from './model';
import { parseSeries } from './parse-series';
import { parseDataLabels, parseTitle } from './parse-text';
import {
	assignDefined,
	attribute,
	cBool,
	cChild,
	cNumber,
	cVal,
	extensionList,
	manualLayout,
	numberFormat,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';

/** Chart group element -> family and whether it is the 3-D variant. */
export const CHART_GROUP_ELEMENTS: Readonly<
	Record<string, { kind: ChartGroupKind; is3D: boolean }>
> = {
	barChart: { kind: 'bar', is3D: false },
	bar3DChart: { kind: 'bar', is3D: true },
	lineChart: { kind: 'line', is3D: false },
	line3DChart: { kind: 'line', is3D: true },
	areaChart: { kind: 'area', is3D: false },
	area3DChart: { kind: 'area', is3D: true },
	pieChart: { kind: 'pie', is3D: false },
	pie3DChart: { kind: 'pie', is3D: true },
	ofPieChart: { kind: 'ofPie', is3D: false },
	doughnutChart: { kind: 'doughnut', is3D: false },
	scatterChart: { kind: 'scatter', is3D: false },
	radarChart: { kind: 'radar', is3D: false },
	bubbleChart: { kind: 'bubble', is3D: false },
	stockChart: { kind: 'stock', is3D: false },
	surfaceChart: { kind: 'surface', is3D: false },
	surface3DChart: { kind: 'surface', is3D: true },
};

const AXIS_ELEMENTS: Readonly<Record<string, ChartAxisKind>> = {
	catAx: 'cat',
	valAx: 'val',
	dateAx: 'date',
	serAx: 'ser',
};

const GROUP = new Set([
	'ser',
	'axId',
	'varyColors',
	'dLbls',
	'barDir',
	'grouping',
	'gapWidth',
	'overlap',
	'gapDepth',
	'shape',
	'marker',
	'dropLines',
	'hiLowLines',
	'upDownBars',
	'serLines',
	'firstSliceAng',
	'holeSize',
	'ofPieType',
	'scatterStyle',
	'radarStyle',
	'bubble3D',
	'bubbleScale',
	'showNegBubbles',
	'sizeRepresents',
	'wireframe',
	'extLst',
]);

function parseGroup(context: ChartParseContext, element: XmlElement): ChartPlotGroup {
	const { kind, is3D } = CHART_GROUP_ELEMENTS[element.localName] ?? { kind: 'bar', is3D: false };
	reportUnmodelled(context, element, GROUP);
	const number = (local: string) => cNumber(context, element, local);
	const flag = (local: string) => cBool(context, element, local);
	const barDir = cVal(element, 'barDir');
	const present = (local: string) => (cChild(element, local) ? true : undefined);
	return assignDefined<ChartPlotGroup>(
		{
			kind,
			element: element.localName,
			is3D,
			series: children(element, 'ser', NS.c).map((ser) => parseSeries(context, ser)),
			axisIds: children(element, 'axId', NS.c).flatMap((axis) => {
				const raw = attribute(axis, 'val')?.trim();
				const id = raw ? Number(raw) : Number.NaN;
				return Number.isFinite(id) ? [id] : [];
			}),
		},
		{
			varyColors: flag('varyColors'),
			dataLabels: parseDataLabels(context, cChild(element, 'dLbls')),
			barDirection: barDir === 'bar' || barDir === 'col' ? barDir : undefined,
			grouping: cVal(element, 'grouping'),
			gapWidth: number('gapWidth'),
			overlap: number('overlap'),
			gapDepth: number('gapDepth'),
			shape: cVal(element, 'shape'),
			marker: flag('marker'),
			dropLines: present('dropLines'),
			hiLowLines: present('hiLowLines'),
			upDownBars: present('upDownBars'),
			seriesLines: present('serLines'),
			firstSliceAngle: number('firstSliceAng'),
			holeSize: number('holeSize'),
			ofPieType: cVal(element, 'ofPieType'),
			scatterStyle: cVal(element, 'scatterStyle'),
			radarStyle: cVal(element, 'radarStyle'),
			bubble3D: flag('bubble3D'),
			bubbleScale: number('bubbleScale'),
			showNegativeBubbles: flag('showNegBubbles'),
			sizeRepresents: cVal(element, 'sizeRepresents'),
			wireframe: flag('wireframe'),
			extLst: extensionList(element),
		},
	);
}

const SCALING = new Set(['orientation', 'logBase', 'min', 'max', 'extLst']);
const AXIS = new Set([
	'axId',
	'scaling',
	'delete',
	'axPos',
	'majorGridlines',
	'minorGridlines',
	'title',
	'numFmt',
	'majorTickMark',
	'minorTickMark',
	'tickLblPos',
	'spPr',
	'txPr',
	'crossAx',
	'crosses',
	'crossesAt',
	'crossBetween',
	'majorUnit',
	'minorUnit',
	'auto',
	'lblAlgn',
	'lblOffset',
	'tickLblSkip',
	'tickMarkSkip',
	'noMultiLvlLbl',
	'baseTimeUnit',
	'extLst',
]);

function parseAxis(
	context: ChartParseContext,
	element: XmlElement,
	kind: ChartAxisKind,
): ChartAxis {
	reportUnmodelled(context, element, AXIS);
	const number = (local: string) => cNumber(context, element, local);
	const scalingElement = cChild(element, 'scaling');
	if (scalingElement) reportUnmodelled(context, scalingElement, SCALING);
	const scaling = assignDefined<ChartAxisScaling>(
		{},
		{
			orientation: cVal(scalingElement, 'orientation'),
			min: cNumber(context, scalingElement, 'min'),
			max: cNumber(context, scalingElement, 'max'),
			logBase: cNumber(context, scalingElement, 'logBase'),
		},
	);
	const title = cChild(element, 'title');
	return assignDefined<ChartAxis>(
		{
			kind,
			scaling,
			majorGridlines: cChild(element, 'majorGridlines') !== undefined,
			minorGridlines: cChild(element, 'minorGridlines') !== undefined,
		},
		{
			id: number('axId'),
			crossAxisId: number('crossAx'),
			position: cVal(element, 'axPos'),
			deleted: cBool(context, element, 'delete'),
			title: title ? parseTitle(context, title) : undefined,
			numberFormat: numberFormat(element),
			majorTickMark: cVal(element, 'majorTickMark'),
			minorTickMark: cVal(element, 'minorTickMark'),
			tickLabelPosition: cVal(element, 'tickLblPos'),
			crosses: cVal(element, 'crosses'),
			crossesAt: number('crossesAt'),
			crossBetween: cVal(element, 'crossBetween'),
			majorUnit: number('majorUnit'),
			minorUnit: number('minorUnit'),
			auto: cBool(context, element, 'auto'),
			labelAlign: cVal(element, 'lblAlgn'),
			labelOffset: number('lblOffset'),
			tickLabelSkip: number('tickLblSkip'),
			tickMarkSkip: number('tickMarkSkip'),
			noMultiLevelLabels: cBool(context, element, 'noMultiLvlLbl'),
			baseTimeUnit: cVal(element, 'baseTimeUnit'),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}

const PLOT_AREA = new Set(['layout', 'spPr', 'extLst']);

/** The plot area (`c:plotArea`): chart groups and axes in document order. */
export function parsePlotArea(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartPlotArea {
	const plotArea: ChartPlotArea = { groups: [], axes: [] };
	if (!element) {
		context.issues.push({
			code: 'CHART_PLOT_AREA_MISSING',
			message: 'The chart has no c:plotArea.',
		});
		return plotArea;
	}
	const handled = new Set(PLOT_AREA);
	for (const child of elements(element)) {
		if (CHART_GROUP_ELEMENTS[child.localName]) {
			plotArea.groups.push(parseGroup(context, child));
			handled.add(child.localName);
		}
		const axis = AXIS_ELEMENTS[child.localName];
		if (axis) {
			plotArea.axes.push(parseAxis(context, child, axis));
			handled.add(child.localName);
		}
	}
	reportUnmodelled(context, element, handled);
	return assignDefined(plotArea, {
		layout: manualLayout(element),
		spPr: shapeProperties(element),
		extLst: extensionList(element),
	});
}
