// Format-neutral model of a DrawingML chart part (`c:chartSpace`, ECMA-376 Part 1, 21.2): the chart
// groups, axes, legend and the part itself. Series, data sources, text and shape properties are in
// `model-series.ts`.
import type {
	ChartDataLabels,
	ChartLayout,
	ChartLines,
	ChartNumberFormat,
	ChartParseIssue,
	ChartShapeProperties,
	ChartSpaceSeries,
	ChartTextBody,
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
	/** Line: `c:smooth` of the group (series carry their own). */
	smooth?: boolean;
	dropLines?: boolean;
	hiLowLines?: boolean;
	/** Line and stock: whether up/down bars are drawn (`c:upDownBars`). */
	upDownBars?: boolean;
	seriesLines?: boolean;
	/** The shapes of `c:dropLines`, `c:hiLowLines` and `c:serLines` when they are written. */
	dropLinesSpPr?: ChartShapeProperties;
	hiLowLinesSpPr?: ChartShapeProperties;
	seriesLinesSpPr?: ChartShapeProperties;
	/** `c:upDownBars/c:gapWidth`. */
	upDownBarsGapWidth?: number;
	/** `c:upBars` and `c:downBars`, present when written. */
	upBars?: ChartLines;
	downBars?: ChartLines;
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
	/** The shapes of the gridlines (`c:majorGridlines/c:spPr`). */
	majorGridlinesSpPr?: ChartShapeProperties;
	minorGridlinesSpPr?: ChartShapeProperties;
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
	txPr?: ChartTextBody;
	extLst?: string;
}

/** The plot area (`c:plotArea`). */
export interface ChartPlotArea {
	/** `c:layout`: `{}` for an empty (automatic) layout element, absent when not written. */
	layout?: ChartLayout;
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
	entries: { index?: number; deleted?: boolean; txPr?: ChartTextBody }[];
	layout?: ChartLayout;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
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
	/** `c14:style` (101..148), the `mc:Choice` Office writes beside the `c:style` fallback. */
	c14Style?: number;
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
	txPr?: ChartTextBody;
	/** Relationship id of the embedded or linked workbook (`c:externalData/@r:id`). */
	externalDataRelId?: string;
	/** `c:externalData/c:autoUpdate`. */
	externalDataAutoUpdate?: boolean;
	/** `c:printSettings` as written (header, footer, margins, page setup), kept for round-trip. */
	printSettings?: string;
	/** Relationship id of the user shapes drawing (`c:userShapes/@r:id`). */
	userShapesRelId?: string;
	/** `c:chart/c:extLst` as written. */
	chartExtLst?: string;
	/** `c:chartSpace/c:extLst` as written. */
	extLst?: string;
	/**
	 * Namespace declarations of the part root besides `c`, `a` and `r` (Office declares `c16r2`),
	 * in document order, so extension lists keep the prefixes they were written with.
	 */
	namespaceDeclarations?: { prefix: string; uri: string }[];
}

/** What {@link parseChartSpace} returns: the model and what it did not model. */
export interface ChartSpaceParseResult {
	chartSpace: ChartSpace;
	issues: ChartParseIssue[];
}
