// Format-neutral model of a series' analysis elements (ECMA-376 Part 1, 21.2): trendlines
// (`c:trendline`, CT_Trendline, with its label `c:trendlineLbl`) and error bars (`c:errBars`,
// CT_ErrBars). The schema's enumerations are held as written; `ChartTrendlineKind` is the readable
// name the fitting code (`trendline-fit.ts`) and the products' editors use.
import type {
	ChartDataSource,
	ChartLayout,
	ChartNumberFormat,
	ChartShapeProperties,
	ChartText,
	ChartTextBody,
} from './model-series';

/** ST_TrendlineType, as written. */
export type ChartTrendlineType = 'exp' | 'linear' | 'log' | 'movingAvg' | 'poly' | 'power';

/** The readable name of a trendline type (`exp` -> `exponential`). */
export type ChartTrendlineKind =
	| 'linear'
	| 'exponential'
	| 'logarithmic'
	| 'polynomial'
	| 'power'
	| 'movingAvg';

const KIND_OF_TYPE: Readonly<Record<ChartTrendlineType, ChartTrendlineKind>> = {
	linear: 'linear',
	exp: 'exponential',
	log: 'logarithmic',
	poly: 'polynomial',
	power: 'power',
	movingAvg: 'movingAvg',
};

const TYPE_OF_KIND: Readonly<Record<ChartTrendlineKind, ChartTrendlineType>> = {
	linear: 'linear',
	exponential: 'exp',
	logarithmic: 'log',
	polynomial: 'poly',
	power: 'power',
	movingAvg: 'movingAvg',
};

/** The readable kind of an ST_TrendlineType value; `undefined` for anything else. */
export const chartTrendlineKind = (type: string | undefined): ChartTrendlineKind | undefined =>
	type !== undefined && Object.hasOwn(KIND_OF_TYPE, type)
		? KIND_OF_TYPE[type as ChartTrendlineType]
		: undefined;

/** The ST_TrendlineType value of a readable kind. */
export const chartTrendlineType = (kind: ChartTrendlineKind): ChartTrendlineType =>
	TYPE_OF_KIND[kind];

/** The label of a trendline (`c:trendlineLbl`): its equation and R-squared text. */
export interface ChartTrendlineLabel {
	layout?: ChartLayout;
	tx?: ChartText;
	numberFormat?: ChartNumberFormat;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	extLst?: string;
}

/** One trendline of a series (`c:trendline`). */
export interface ChartTrendline {
	/** `c:name`, as written. */
	name?: string;
	spPr?: ChartShapeProperties;
	/** `c:trendlineType/@val`; absent when missing or not an ST_TrendlineType value. */
	trendlineType?: ChartTrendlineType;
	/** Polynomial order (2..6). */
	order?: number;
	/** Moving-average period (2 or more). */
	period?: number;
	/** Extension of the line beyond the data, in category or x units. */
	forward?: number;
	backward?: number;
	/** The y value the line is forced through at x = 0. */
	intercept?: number;
	displayRSquared?: boolean;
	displayEquation?: boolean;
	label?: ChartTrendlineLabel;
	extLst?: string;
}

/** ST_ErrDir. */
export type ChartErrorBarDirection = 'x' | 'y';
/** ST_ErrBarType: both sides, the negative side or the positive side. */
export type ChartErrorBarType = 'both' | 'minus' | 'plus';
/** ST_ErrValType: how the error amount is calculated. */
export type ChartErrorValueType = 'cust' | 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr';

/** One set of error bars of a series (`c:errBars`). */
export interface ChartErrorBars {
	/** `c:errDir`; absent when not written (y for most charts). */
	direction?: ChartErrorBarDirection;
	/** `c:errBarType`; absent when missing or unreadable. */
	barType?: ChartErrorBarType;
	/** `c:errValType`; absent when missing or unreadable. */
	valueType?: ChartErrorValueType;
	noEndCap?: boolean;
	/** Custom amounts (`c:plus`, `c:minus`), CT_NumDataSource. */
	plus?: ChartDataSource;
	minus?: ChartDataSource;
	/** `c:val`: the fixed value, percentage or standard-deviation multiple. */
	value?: number;
	spPr?: ChartShapeProperties;
	extLst?: string;
}

export const CHART_ERROR_BAR_DIRECTIONS: readonly ChartErrorBarDirection[] = ['x', 'y'];
export const CHART_ERROR_BAR_TYPES: readonly ChartErrorBarType[] = ['both', 'minus', 'plus'];
export const CHART_ERROR_VALUE_TYPES: readonly ChartErrorValueType[] = [
	'cust',
	'fixedVal',
	'percentage',
	'stdDev',
	'stdErr',
];
