// Format-neutral model of an Office 2016 extended chart part (`cx:chartSpace`, [MS-ODRAWXML]
// 2.24.3, the chartex namespace): waterfall, histogram, pareto, box and whisker, treemap, sunburst,
// funnel and region map charts. One model for PowerPoint, Word, Excel and Visio. Values are kept as
// written (cached point text, raw enumerations) so clients choose their own defaults; shape and text
// properties go through the same `drawingml` models as the classic `c:chartSpace` model; pieces the
// model does not break down (geography caches, colour-by-value scales, format overrides, print
// settings, extension lists) stay raw XML for round-trip.
import type {
	ChartCachePoint,
	ChartLines,
	ChartNumberFormat,
	ChartParseIssue,
	ChartShapeProperties,
	ChartTextBody,
} from '../model-series';

/** A formula (`cx:f`, `cx:nf`). */
export interface ChartExFormula {
	/** The formula text as written (`Sheet1!$B$2:$B$5`, or Excel's `_xlchart.v1.7`). */
	formula: string;
	/** `@dir`: how a two-dimensional range is read (`row` or `col`). */
	direction?: string;
}

/** One level of cached points of a dimension (`cx:lvl`). */
export interface ChartExLevel {
	/** `@ptCount`; absent when not written. */
	pointCount?: number;
	/** `@formatCode` of a numeric level. */
	formatCode?: string;
	/** `@name`. */
	name?: string;
	/** Cached points (`cx:pt`), text exactly as written. */
	points: ChartCachePoint[];
	extLst?: string;
}

/** A data dimension (`cx:numDim`, `cx:strDim`). */
export interface ChartExDimension {
	/** `numDim` holds numbers, `strDim` text. */
	kind: 'numDim' | 'strDim';
	/** `@type`: `val`, `x`, `y`, `size`, `colorVal` (numeric); `cat`, `colorStr`, `entityId` (text). */
	type?: string;
	/** `cx:f`. */
	formula?: ChartExFormula;
	/** `cx:nf`: the formula of the dimension's name. */
	nameFormula?: ChartExFormula;
	/** Levels in document order; a hierarchical category dimension has the leaf level first. */
	levels: ChartExLevel[];
	extLst?: string;
}

/** One data set (`cx:data`), referenced by series through `cx:dataId`. */
export interface ChartExData {
	/** `@id`. */
	id?: number;
	/** Dimensions in document order. */
	dimensions: ChartExDimension[];
	extLst?: string;
}

/** `cx:chartData`. */
export interface ChartExChartData {
	/** `cx:externalData`: the embedded or linked workbook. */
	externalData?: { relId?: string; autoUpdate?: boolean };
	data: ChartExData[];
}

/** Chart text (`cx:tx`): a cell reference with its cached text (`cx:txData`), or rich text. */
export interface ChartExText {
	/** `cx:txData/cx:f`. */
	formula?: ChartExFormula;
	/** `cx:txData/cx:v`, as written. */
	value?: string;
	/** `cx:rich`. */
	rich?: ChartTextBody;
}

/** A chart, axis or units title (`cx:title`, `cx:unitsLabel`). */
export interface ChartExTitle {
	tx?: ChartExText;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	/** Chart title: `@pos` (`t`, `b`, `l`, `r`). */
	position?: string;
	/** Chart title: `@align` (`min`, `ctr`, `max`). */
	align?: string;
	/** Chart title: `@overlay`. */
	overlay?: boolean;
	extLst?: string;
}

/** `cx:legend`. */
export interface ChartExLegend {
	/** `@pos` (`t`, `b`, `l`, `r`). */
	position?: string;
	/** `@align` (`min`, `ctr`, `max`). */
	align?: string;
	overlay?: boolean;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	extLst?: string;
}

/** What a data label shows (`cx:visibility` of `cx:dataLabels` and `cx:dataLabel`). */
export interface ChartExLabelVisibility {
	seriesName?: boolean;
	categoryName?: boolean;
	value?: boolean;
}

/** Settings shared by `cx:dataLabels` and `cx:dataLabel`. */
export interface ChartExDataLabelOptions {
	/** `@pos` (`bestFit`, `b`, `ctr`, `inBase`, `inEnd`, `l`, `outEnd`, `r`, `t`). */
	position?: string;
	numberFormat?: ChartNumberFormat;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	visibility?: ChartExLabelVisibility;
	/** `cx:separator` text. */
	separator?: string;
	extLst?: string;
}

/** One point's label override (`cx:dataLabel`). */
export interface ChartExDataLabel extends ChartExDataLabelOptions {
	index?: number;
}

/** The data labels of a series (`cx:dataLabels`). */
export interface ChartExDataLabels extends ChartExDataLabelOptions {
	labels: ChartExDataLabel[];
	/** `cx:dataLabelHidden/@idx` values. */
	hidden: number[];
}

/** A data point override (`cx:dataPt`). */
export interface ChartExDataPoint {
	index?: number;
	spPr?: ChartShapeProperties;
	extLst?: string;
}

/** A number or `auto`, as the chartex schema writes optional doubles. */
export type ChartExAutoNumber = number | 'auto';

/** Histogram binning (`cx:binning`). */
export interface ChartExBinning {
	/** `@intervalClosed`: `l` or `r`. */
	intervalClosed?: string;
	underflow?: ChartExAutoNumber;
	overflow?: ChartExAutoNumber;
	/** `cx:binSize/@val`. */
	binSize?: number;
	/** `cx:binCount/@val`. */
	binCount?: number;
}

/** What a series shows of its parts (`cx:layoutPr/cx:visibility`). */
export interface ChartExSeriesVisibility {
	connectorLines?: boolean;
	meanLine?: boolean;
	meanMarker?: boolean;
	nonoutliers?: boolean;
	outliers?: boolean;
}

/** Region map geography (`cx:geography`): the attributes, and the provider cache as written. */
export interface ChartExGeography {
	projectionType?: string;
	viewedRegionType?: string;
	cultureLanguage?: string;
	cultureRegion?: string;
	attribution?: string;
	/** `cx:geoCache` as written. */
	geoCacheXml?: string;
	extLst?: string;
}

/** Family-specific series options (`cx:layoutPr`). */
export interface ChartExLayoutProperties {
	/** Treemap: `none`, `banner`, `overlapping`. */
	parentLabelLayout?: string;
	/** Region map: `none`, `bestFitOnly`, `showAll`. */
	regionLabelLayout?: string;
	visibility?: ChartExSeriesVisibility;
	/** Histogram: bins by category (`cx:aggregation`) instead of `binning`. */
	aggregation?: boolean;
	binning?: ChartExBinning;
	geography?: ChartExGeography;
	/** Box and whisker: `cx:statistics`, with `@quartileMethod` (`inclusive`, `exclusive`). */
	statistics?: { quartileMethod?: string };
	/** Waterfall: point indexes drawn as totals (`cx:subtotals/cx:idx`). */
	subtotals?: number[];
	extLst?: string;
}

/** One series (`cx:series`). */
export interface ChartExSeries {
	/** `@layoutId`: `waterfall`, `clusteredColumn`, `paretoLine`, `boxWhisker`, `treemap`... */
	layoutId: string;
	hidden?: boolean;
	/** `@ownerIdx`: the series a pareto line belongs to. */
	ownerIndex?: number;
	uniqueId?: string;
	formatIndex?: number;
	/** Series name (`cx:tx`). */
	tx?: ChartExText;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	/** `cx:valueColors` as written (region map colour scale). */
	valueColorsXml?: string;
	/** `cx:valueColorPositions` as written. */
	valueColorPositionsXml?: string;
	dataPoints: ChartExDataPoint[];
	dataLabels?: ChartExDataLabels;
	/**
	 * A `cx:data` written inside the series instead of referenced through `cx:dataId`. Not in the
	 * schema; read (and written back) because some producers write it and PowerPoint reads it.
	 */
	data?: ChartExData;
	/** `cx:dataId/@val`: the `cx:data` this series draws. */
	dataId?: number;
	/** `cx:layoutPr`; `{}` for an empty element. */
	layoutProperties?: ChartExLayoutProperties;
	/** `cx:axisId/@val` values. */
	axisIds: number[];
	extLst?: string;
}

/** Axis scaling: `cx:catScaling` or `cx:valScaling`. */
export type ChartExAxisScaling =
	| { kind: 'cat'; gapWidth?: ChartExAutoNumber }
	| {
			kind: 'val';
			max?: ChartExAutoNumber;
			min?: ChartExAutoNumber;
			majorUnit?: ChartExAutoNumber;
			minorUnit?: ChartExAutoNumber;
	  };

/** Display units of a value axis (`cx:units`). */
export interface ChartExAxisUnits {
	/** `@unit`: `hundreds`, `thousands`... `trillions`, `percentage`. */
	unit?: string;
	label?: ChartExTitle;
	extLst?: string;
}

/** One axis (`cx:axis`). */
export interface ChartExAxis {
	id?: number;
	hidden?: boolean;
	/** Absent when the axis has neither scaling element. */
	scaling?: ChartExAxisScaling;
	title?: ChartExTitle;
	units?: ChartExAxisUnits;
	majorGridlines?: ChartLines;
	minorGridlines?: ChartLines;
	/** `cx:majorTickMarks`, with its `@type` (`in`, `out`, `cross`, `none`) when written. */
	majorTickMarks?: { type?: string };
	minorTickMarks?: { type?: string };
	/** `cx:tickLabels`: present when tick labels are shown. */
	tickLabels?: { extLst?: string };
	numberFormat?: ChartNumberFormat;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	extLst?: string;
}

/** `cx:plotArea`. */
export interface ChartExPlotArea {
	/** `cx:plotAreaRegion/cx:plotSurface`. */
	plotSurface?: { spPr?: ChartShapeProperties; extLst?: string };
	series: ChartExSeries[];
	/** `cx:plotAreaRegion/cx:extLst`. */
	regionExtLst?: string;
	axes: ChartExAxis[];
	spPr?: ChartShapeProperties;
	extLst?: string;
}

/** An extended chart part (`cx:chartSpace`). */
export interface ChartExSpace {
	chartData: ChartExChartData;
	title?: ChartExTitle;
	plotArea: ChartExPlotArea;
	legend?: ChartExLegend;
	/** `cx:chart/cx:extLst`. */
	chartExtLst?: string;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	/** `cx:clrMapOvr`, `cx:fmtOvrs` and `cx:printSettings` as written. */
	colorMapOverrideXml?: string;
	formatOverridesXml?: string;
	printSettingsXml?: string;
	/**
	 * Relationship id of the user shapes drawing. PowerPoint writes the classic `c:userShapes`
	 * element under the chartex root for it.
	 */
	userShapesRelId?: string;
	extLst?: string;
	/** Namespace declarations of the root in document order; defaults to `cx`, `a`, `r`. */
	namespaceDeclarations?: { prefix: string; uri: string }[];
}

/** What {@link parseChartExSpace} returns: the model and what it did not model. */
export interface ChartExSpaceParseResult {
	chartExSpace: ChartExSpace;
	issues: ChartParseIssue[];
}
