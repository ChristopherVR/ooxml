/**
 * Structural fields of pptx chart markers, data points and data labels, filled from the neutral
 * chart model onto the objects the object-tree readers built (which keep the resolved styles):
 * marker symbol and size, point index, explosion, invert-if-negative and 3-D bubble, and the
 * data-label switches, position, separator and number format. A field the neutral model leaves
 * unset is removed, so the neutral model alone decides it.
 *
 * @module chart-from-neutral-labels
 */
import type {
	ChartDataLabel,
	ChartDataLabelOptions,
	ChartDataLabels,
	ChartDataPoint,
	ChartMarker,
} from '../../../chart/index';
import type {
	PptxChartDataLabel,
	PptxChartDataLabelOptions,
	PptxChartDataLabelPosition,
	PptxChartDataPoint,
	PptxChartMarker,
	PptxChartMarkerSymbol,
} from '../types';

/** Sets `key` to `value`, or removes it when `value` is undefined. */
export function setOrRemove<T extends object, K extends keyof T>(
	target: T,
	key: K,
	value: T[K] | undefined,
): void {
	if (value === undefined) delete target[key];
	else target[key] = value;
}

const MARKER_SYMBOLS = new Set<string>([
	'circle',
	'dash',
	'diamond',
	'dot',
	'none',
	'picture',
	'plus',
	'square',
	'star',
	'triangle',
	'x',
	'auto',
]);

/**
 * A marker: an unknown or missing `c:symbol` is `auto` (PowerPoint still draws an automatic
 * marker), and a size outside ST_MarkerSize (2..72) is dropped. `styled` is the object-tree marker
 * holding the resolved `c:spPr`.
 */
export function markerFromNeutral(
	marker: ChartMarker | undefined,
	styled: PptxChartMarker | undefined,
): PptxChartMarker | undefined {
	if (!marker) return styled;
	const symbol = marker.symbol?.trim() ?? '';
	const out: PptxChartMarker = styled ?? { symbol: 'auto' };
	out.symbol = (MARKER_SYMBOLS.has(symbol) ? symbol : 'auto') as PptxChartMarkerSymbol;
	const size = marker.size;
	setOrRemove(
		out,
		'size',
		size !== undefined && Number.isInteger(size) && size >= 2 && size <= 72 ? size : undefined,
	);
	return out;
}

const unsignedInt = (value: number | undefined) =>
	value !== undefined && Number.isInteger(value) && value >= 0 && value <= 0xffffffff
		? value
		: undefined;

/** Takes the first entry of `pool` whose index is `index` out of the pool. */
function take<T extends { index?: number }>(pool: T[], index: number): T | undefined {
	const at = pool.findIndex((entry) => entry.index === index);
	return at < 0 ? undefined : pool.splice(at, 1)[0];
}

/** Data point overrides (`c:dPt`), matched to the object-tree points by index. */
export function dataPointsFromNeutral(
	points: ChartDataPoint[],
	styled: PptxChartDataPoint[] | undefined,
): PptxChartDataPoint[] | undefined {
	if (!styled) return undefined;
	const pool = [...points];
	for (const point of styled) {
		const neutral = take(pool, point.idx);
		if (!neutral) continue;
		setOrRemove(point, 'explosion', unsignedInt(neutral.explosion));
		setOrRemove(point, 'invertIfNegative', neutral.invertIfNegative);
		setOrRemove(point, 'bubble3D', neutral.bubble3D);
		setOrRemove(point, 'marker', markerFromNeutral(neutral.marker, point.marker));
	}
	return styled;
}

const POSITIONS = new Set<string>([
	'bestFit',
	'b',
	'ctr',
	'inBase',
	'inEnd',
	'l',
	'outEnd',
	'r',
	't',
]);

const position = (value: string | undefined) =>
	value !== undefined && POSITIONS.has(value) ? (value as PptxChartDataLabelPosition) : undefined;

const formatCode = (options: ChartDataLabelOptions) =>
	options.numberFormat?.formatCode.trim() || undefined;

/** Group or series data-label options (`c:dLbls`). */
export function labelOptionsFromNeutral(
	labels: ChartDataLabels | undefined,
	styled: PptxChartDataLabelOptions | undefined,
): PptxChartDataLabelOptions | undefined {
	if (!labels) return styled;
	const out: PptxChartDataLabelOptions = styled ?? {};
	setOrRemove(out, 'showValue', labels.showValue);
	setOrRemove(out, 'showCategory', labels.showCategoryName);
	setOrRemove(out, 'showSeriesName', labels.showSeriesName);
	setOrRemove(out, 'showPercent', labels.showPercent);
	setOrRemove(out, 'showLegendKey', labels.showLegendKey);
	setOrRemove(out, 'showBubbleSize', labels.showBubbleSize);
	setOrRemove(out, 'showLeaderLines', labels.showLeaderLines);
	setOrRemove(out, 'position', position(labels.position));
	setOrRemove(out, 'separator', labels.separator);
	setOrRemove(out, 'numberFormat', formatCode(labels));
	return out;
}

/** One point's label (`c:dLbl`); its text and layout stay with the object-tree reader. */
function pointLabelFromNeutral(label: ChartDataLabel, out: PptxChartDataLabel): void {
	setOrRemove(out, 'deleted', label.deleted);
	setOrRemove(out, 'showVal', label.showValue);
	setOrRemove(out, 'showCatName', label.showCategoryName);
	setOrRemove(out, 'showSerName', label.showSeriesName);
	setOrRemove(out, 'showPercent', label.showPercent);
	setOrRemove(out, 'showLegendKey', label.showLegendKey);
	setOrRemove(out, 'showBubbleSize', label.showBubbleSize);
	setOrRemove(out, 'position', position(label.position));
	setOrRemove(out, 'separator', label.separator);
	setOrRemove(out, 'numberFormat', formatCode(label));
}

/** Point labels (`c:dLbls/c:dLbl`), matched to the object-tree labels by index. */
export function pointLabelsFromNeutral(
	labels: ChartDataLabels | undefined,
	styled: PptxChartDataLabel[] | undefined,
): PptxChartDataLabel[] | undefined {
	if (!styled || !labels) return styled;
	const pool = [...labels.labels];
	for (const label of styled) {
		const neutral = take(pool, label.idx);
		if (neutral) pointLabelFromNeutral(neutral, label);
	}
	return styled;
}
