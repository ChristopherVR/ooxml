/** Numeric series required by the shared axis algorithms. */
export interface ChartAxisSeries {
	values: readonly (number | null)[];
}
/** Axis scaling inputs; adapters map native chart models onto these fields. */
export interface ChartAxisOptions {
	axisType?: string;
	axPos?: string;
	min?: number;
	max?: number;
	logScale?: boolean;
	logBase?: number;
	orientation?: string;
	majorUnit?: number;
	minorUnit?: number;
}
/** Data-space bounds, with a logarithmic span when requested. */
export interface ChartAxisRange {
	min: number;
	max: number;
	span: number;
	logScale?: boolean;
	logBase?: number;
	reverseOrder?: boolean;
	majorUnit?: number;
}
