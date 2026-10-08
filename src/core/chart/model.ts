// Format-neutral model of a DrawingML chart part (`c:chartSpace`, ECMA-376 Part 1, 21.2): the chart
// groups, axes, legend and the part itself. Series, data sources, text and shape properties are in
// `model-series.ts`.
import type { DrawingTextBody } from '../drawingml/types';
import type { ChartManualLayout } from './manual-layout';
import type {
	ChartDataLabels,
	ChartNumberFormat,
	ChartParseIssue,
	ChartShapeProperties,
	ChartSpaceSeries,
	ChartTitle,
} from './model-series';

export type * from './model-series';

/** The chart family of a plot group. */
export type ChartGroupKind =
	| 'bar'
	| 'line'
	| 'area'
	| 'pie'
	| 'doughnut'
	| 'ofPie'
	| 'scatter'
	| 'radar'
	| 'bubble'
	| 'stock'
	| 'surface';

/** One chart group of the plot area (`c:barChart`, `c:pie3DChart`, `c:stockChart`...). */
export interface ChartPlotGroup {
	kind: ChartGroupKind;
	/** The element's local name, e.g. `bar3DChart`. */
	element: string;
	is3D: boolean;
	series: ChartSpaceSeries[];
	/** `c:axId` values, in order. */
	axisIds: number[];
	varyColors?: boolean;
	dataLabels?: ChartDataLabels;
	/** Bar: `bar` (horizontal) or `col` (vertical). */
	barDirection?: 'bar' | 'col';
	/** Bar, line and area: `clustered`, `stacked`, `percentStacked`, `standard`. */
	grouping?: string;
	gapWidth?: number;
	overlap?: number;
	gapDepth?: number;
	/** 3-D bar shape (`box`, `cylinder`...). */
	shape?: string;
	/** Line: whether markers are shown (`c:marker`). */
	marker?: boolean;
	dropLines?: boolean;
	hiLowLines?: boolean;
	/** Line and stock: whether up/down bars are drawn (`c:upDownBars`). */
	upDownBars?: boolean;
	seriesLines?: boolean;
	/** Pie and doughnut, degrees. */
	firstSliceAngle?: number;
	/** Doughnut, percent. */
	holeSize?: number;
	/** Pie of pie / bar of pie: `pie` or `bar`. */
	ofPieType?: string;
	/** Scatter: `lineMarker`, `smoothMarker`, `marker`... */
	scatterStyle?: string;
	/** Radar: `standard`, `marker`, `filled`. */
	radarStyle?: string;
	bubble3D?: boolean;
	bubbleScale?: number;
	showNegativeBubbles?: boolean;
	/** Bubble: `area` or `w`. */
	sizeRepresents?: string;
	/** Surface: drawn as a wireframe. */
	wireframe?: boolean;
	extLst?: string;
}

/** The axis families. */
export type ChartAxisKind = 'cat' | 'val' | 'date' | 'ser';

/** Axis scaling (`c:scaling`). */
export interface ChartAxisScaling {
	/** `minMax` (normal) or `maxMin` (reversed). */
	orientation?: string;
	min?: number;
	max?: number;
	logBase?: number;
}

/** One axis (`c:catAx`, `c:valAx`, `c:dateAx`, `c:serAx`). */
export interface ChartAxis {
	kind: ChartAxisKind;
	id?: number;
	crossAxisId?: number;
	/** `b`, `l`, `r`, `t`. */
	position?: string;
	deleted?: boolean;
	scaling: ChartAxisScaling;
	majorGridlines: boolean;
	minorGridlines: boolean;
	title?: ChartTitle;
	numberFormat?: ChartNumberFormat;
	/** `cross`, `in`, `out`, `none`. */
	majorTickMark?: string;
	minorTickMark?: string;
	/** `nextTo`, `low`, `high`, `none`. */
	tickLabelPosition?: string;
	/** `autoZero`, `min`, `max`. */
	crosses?: string;
	crossesAt?: number;
	/** Value axis: `between` or `midCat`. */
	crossBetween?: string;
	majorUnit?: number;
	minorUnit?: number;
	auto?: boolean;
	labelAlign?: string;
	labelOffset?: number;
	tickLabelSkip?: number;
	tickMarkSkip?: number;
	noMultiLevelLabels?: boolean;
	baseTimeUnit?: string;
	spPr?: ChartShapeProperties;
	txPr?: DrawingTextBody;
	extLst?: string;
}

/** The plot area (`c:plotArea`). */
export interface ChartPlotArea {
	layout?: ChartManualLayout;
	/** Chart groups in document order; a combination chart has more than one. */
	groups: ChartPlotGroup[];
	axes: ChartAxis[];
	spPr?: ChartShapeProperties;
	extLst?: string;
}

/** The legend (`c:legend`). */
export interface ChartLegend {
	/** `r`, `l`, `t`, `b`, `tr`. */
	position?: string;
	overlay?: boolean;
	/** `c:legendEntry` overrides. */
	entries: { index?: number; deleted?: boolean; txPr?: DrawingTextBody }[];
	layout?: ChartManualLayout;
	spPr?: ChartShapeProperties;
	txPr?: DrawingTextBody;
	extLst?: string;
}

/** 3-D view (`c:view3D`). */
export interface ChartView3D {
	rotX?: number;
	rotY?: number;
	depthPercent?: number;
	heightPercent?: number;
	rightAngleAxes?: boolean;
	perspective?: number;
}

/** A chart part (`c:chartSpace`). */
export interface ChartSpace {
	date1904?: boolean;
	language?: string;
	roundedCorners?: boolean;
	/** `c:style` (1..48), read through `mc:AlternateContent` when Office wraps it. */
	style?: number;
	title?: ChartTitle;
	autoTitleDeleted?: boolean;
	view3D?: ChartView3D;
	plotArea: ChartPlotArea;
	legend?: ChartLegend;
	plotVisibleOnly?: boolean;
	/** `gap`, `span`, `zero`. */
	displayBlanksAs?: string;
	showDataLabelsOverMax?: boolean;
	spPr?: ChartShapeProperties;
	txPr?: DrawingTextBody;
	/** Relationship id of the embedded or linked workbook (`c:externalData/@r:id`). */
	externalDataRelId?: string;
	/** Relationship id of the user shapes drawing (`c:userShapes/@r:id`). */
	userShapesRelId?: string;
	/** `c:chart/c:extLst` as written. */
	chartExtLst?: string;
	/** `c:chartSpace/c:extLst` as written. */
	extLst?: string;
}

/** What {@link parseChartSpace} returns: the model and what it did not model. */
export interface ChartSpaceParseResult {
	chartSpace: ChartSpace;
	issues: ChartParseIssue[];
}
