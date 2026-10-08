/**
 * The comparison shape of the chart parity harness (`chart-neutral-parity.test.ts`): the pptx chart model
 * (`PptxChartData`, parsed from the `fast-xml-parser` object tree) and the neutral chart model
 * (`ChartSpace`, parsed by `parseChartSpace` over the shared DOM) are each reduced to one comparison
 * shape, {@link ChartParityShape}, so a chart part read by both parsers can be compared field by field.
 *
 * Only facts BOTH models carry are projected; what one model carries and the other does not is
 * listed in docs/agnostic-core-plan.md (step 5). The neutral model keeps values as written, so the
 * pptx normalisations are applied on the neutral side (`chart-neutral-parity-neutral*.ts`), each one
 * named and justified where it is applied (search for "pptx normalisation"). Representation differences that are not a normalisation (an axis list
 * in a different order, a boolean the pptx model only stores when true) are marked "tolerance".
 *
 * @module __tests__/integration/chart-neutral-parity
 */

/** Data-label switches and position, as both models carry them. */
export interface LabelShape {
	showValue?: boolean;
	showCategory?: boolean;
	showSeriesName?: boolean;
	showPercent?: boolean;
	showLegendKey?: boolean;
	showBubbleSize?: boolean;
	showLeaderLines?: boolean;
	position?: string;
	separator?: string;
	numberFormat?: string;
}

export interface PointLabelShape extends LabelShape {
	idx: number;
	deleted?: boolean;
	text?: string;
}

export interface MarkerShape {
	symbol: string;
	size?: number;
}

export interface SeriesShape {
	chartType: string;
	idx?: number;
	name: string;
	values: (number | null)[];
	xValues?: number[];
	bubbleSizes?: number[];
	formatCode?: string;
	smooth?: boolean;
	invertIfNegative?: boolean;
	explosion?: number;
	shape?: string;
	marker?: MarkerShape;
	dataPoints?: {
		idx: number;
		explosion?: number;
		invertIfNegative?: boolean;
		bubble3D?: boolean;
		marker?: MarkerShape;
	}[];
	labels?: LabelShape;
	pointLabels?: PointLabelShape[];
}

export interface AxisShape {
	kind: string;
	id?: number;
	crossAxisId?: number;
	position?: string;
	deleted?: boolean;
	crosses?: string;
	crossesAt?: number;
	crossBetween?: string;
	orientation?: string;
	min?: number;
	max?: number;
	logBase?: number;
	majorUnit?: number;
	minorUnit?: number;
	majorGridlines?: boolean;
	minorGridlines?: boolean;
	majorTickMark?: string;
	minorTickMark?: string;
	tickLabelPosition?: string;
	numberFormat?: { formatCode: string; sourceLinked: boolean };
	title?: string;
}

/** The comparison shape of one chart part. */
export interface ChartParityShape {
	chartType: string;
	grouping?: string;
	groupingStandard?: boolean;
	barDirection?: string;
	gapWidth?: number;
	overlap?: number;
	gapDepth?: number;
	varyColors?: boolean;
	firstSliceAngle?: number;
	holeSize?: number;
	scatterStyle?: string;
	radarStyle?: string;
	barShape?: string;
	wireframe?: boolean;
	surfaceTopView?: boolean;
	categories: string[];
	categoryLevels?: string[][];
	series: SeriesShape[];
	axes: AxisShape[];
	title?: string;
	legend?: { position?: string; overlay?: boolean };
	groupLabels?: LabelShape;
	plotVisibleOnly?: boolean;
	dispBlanksAs?: string;
	autoTitleDeleted?: boolean;
	style?: number;
	view3D?: Record<string, number | boolean | undefined>;
	date1904?: boolean;
	roundedCorners?: boolean;
}

/** Drops `undefined` entries so `toStrictEqual` compares only what is set. */
export function compact<T extends object>(value: T): T {
	for (const key of Object.keys(value) as (keyof T)[])
		if (value[key] === undefined) delete value[key];
	return value;
}

export const nonEmpty = <T>(list: T[] | undefined): T[] | undefined =>
	list && list.length > 0 ? list : undefined;

/** Tolerance: the pptx axis list is grouped by element name, not in document order. */
export const byAxisId = (left: AxisShape, right: AxisShape) =>
	(left.id ?? -1) - (right.id ?? -1) || left.kind.localeCompare(right.kind);

/** Every path at which two projections differ, as `path: neutral <a> / pptx <b>`. */
export function diffShapes(neutral: unknown, pptx: unknown, at = ''): string[] {
	if (Object.is(neutral, pptx)) return [];
	const isObject = (value: unknown): value is Record<string, unknown> =>
		typeof value === 'object' && value !== null;
	if (isObject(neutral) && isObject(pptx) && Array.isArray(neutral) === Array.isArray(pptx)) {
		const keys = new Set([...Object.keys(neutral), ...Object.keys(pptx)]);
		return [...keys].flatMap((key) => diffShapes(neutral[key], pptx[key], `${at}.${key}`));
	}
	return [`${at || '.'}: neutral ${JSON.stringify(neutral)} / pptx ${JSON.stringify(pptx)}`];
}
