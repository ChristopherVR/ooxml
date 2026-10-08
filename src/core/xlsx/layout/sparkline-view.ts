// Lays out one sparkline for painting: DOM-free geometry in a unit box (x 0..1 left to right,
// y 0..1 top to bottom) with every colour resolved to `#RRGGBB`. The painter only scales the box
// to the host cell. Point colours stack in this order, the later winning: markers, negative,
// high, low, first, last.
import type { Color, SparklineGroup, ThemePalette } from '../model';
import { resolveColor } from './colors';

/** A point of a line sparkline in the unit box. */
export interface SparklinePointView {
	x: number;
	y: number;
}

export interface SparklineMarkerView extends SparklinePointView {
	color: string;
}

export interface SparklineColumnView {
	x: number;
	y: number;
	w: number;
	h: number;
	color: string;
}

export interface SparklineView {
	type: SparklineGroup['type'];
	/** Series colour. */
	color: string;
	/** Line width in points (lines only). */
	lineWeightPt: number;
	/** Polylines of a line sparkline; an empty cell shown as a gap splits them. */
	lines: SparklinePointView[][];
	markers: SparklineMarkerView[];
	columns: SparklineColumnView[];
	/** The horizontal axis at value zero, when shown and inside the scale. */
	axis?: { y: number; color: string };
}

/** The values of one sparkline (null for an empty or non-numeric cell) and optional dates. */
export interface SparklineData {
	values: readonly (number | null)[];
	/** Date serials, one per value, for a date axis. */
	dates?: readonly (number | null)[];
}

export interface SparklineScale {
	min: number;
	max: number;
}

type ColorKey =
	| 'colorSeries'
	| 'colorNegative'
	| 'colorAxis'
	| 'colorMarkers'
	| 'colorFirst'
	| 'colorLast'
	| 'colorHigh'
	| 'colorLow';

/** Excel's default sparkline colours (the first style of its gallery). */
const DEFAULTS: Record<ColorKey, string> = {
	colorSeries: '#376092',
	colorNegative: '#D00000',
	colorAxis: '#000000',
	colorMarkers: '#D00000',
	colorFirst: '#D00000',
	colorLast: '#D00000',
	colorHigh: '#D00000',
	colorLow: '#D00000',
};

/** Fraction of each column slot left empty between columns. */
const COLUMN_GAP = 0.2;

function extent(values: readonly (number | null)[]): SparklineScale | undefined {
	let min = Infinity;
	let max = -Infinity;
	for (const v of values) {
		if (v === null) continue;
		if (v < min) min = v;
		if (v > max) max = v;
	}
	return min <= max ? { min, max } : undefined;
}

/** Values after the group's empty-cell rule (`zero` turns empty cells into 0). */
export function sparklineValues(
	group: SparklineGroup,
	values: readonly (number | null)[],
): (number | null)[] {
	return group.displayEmptyCellsAs === 'zero' ? values.map((v) => v ?? 0) : [...values];
}

/**
 * The vertical scale of the sparkline at `index` of `all` (the group's sparklines in order): its
 * own extent, the group's, or the manual limit, chosen per axis end.
 */
export function sparklineScale(
	group: SparklineGroup,
	all: readonly SparklineData[],
	index: number,
): SparklineScale | undefined {
	const own = extent(sparklineValues(group, all[index]?.values ?? []));
	const needsGroup = group.minAxisType === 'group' || group.maxAxisType === 'group';
	const shared = needsGroup
		? extent(all.flatMap((data) => sparklineValues(group, data.values)))
		: undefined;
	const end = (
		type: SparklineGroup['minAxisType'],
		manual: number | undefined,
		key: 'min' | 'max',
	): number | undefined => {
		if (type === 'custom' && manual !== undefined) return manual;
		if (type === 'group' && shared) return shared[key];
		return own?.[key];
	};
	const min = end(group.minAxisType, group.manualMin, 'min');
	const max = end(group.maxAxisType, group.manualMax, 'max');
	return min === undefined || max === undefined ? undefined : { min, max };
}

/** Horizontal positions (0..1) of each value: evenly spaced, or by date on a date axis. */
function positions(group: SparklineGroup, data: SparklineData, n: number): (number | null)[] {
	const dates = data.dates;
	if (group.dateAxis && dates && dates.length === n) {
		const span = extent(dates);
		if (span) {
			const width = span.max - span.min;
			return dates.map((d) => (d === null ? null : width ? (d - span.min) / width : 0.5));
		}
	}
	return Array.from({ length: n }, (_, i) => (n === 1 ? 0.5 : i / (n - 1)));
}

interface PointFacts {
	high: number;
	low: number;
	first: number;
	last: number;
}

function pointFacts(values: readonly (number | null)[]): PointFacts {
	const numbers = values.filter((x): x is number => x !== null);
	let last = values.length - 1;
	while (last >= 0 && values[last] === null) last--;
	return {
		high: Math.max(...numbers),
		low: Math.min(...numbers),
		first: values.findIndex((x) => x !== null),
		last,
	};
}

/** The colour the group's point flags give the value at `i`, if any. */
function pointColor(
	group: SparklineGroup,
	v: number,
	i: number,
	facts: PointFacts,
	colors: Record<ColorKey, string>,
	markers: boolean,
): string | undefined {
	let color = markers && group.markers ? colors.colorMarkers : undefined;
	if (group.negative && v < 0) color = colors.colorNegative;
	if (group.high && v === facts.high) color = colors.colorHigh;
	if (group.low && v === facts.low) color = colors.colorLow;
	if (group.first && i === facts.first) color = colors.colorFirst;
	if (group.last && i === facts.last) color = colors.colorLast;
	return color;
}

function resolveColors(group: SparklineGroup, theme: ThemePalette): Record<ColorKey, string> {
	const out = { ...DEFAULTS };
	for (const key of Object.keys(DEFAULTS) as ColorKey[]) {
		const color: Color | undefined = group[key];
		out[key] = resolveColor(color, theme, DEFAULTS[key]) ?? DEFAULTS[key];
	}
	return out;
}

/** Lays out one sparkline; `scale` comes from {@link sparklineScale}. */
export function sparklineView(
	group: SparklineGroup,
	data: SparklineData,
	scale: SparklineScale | undefined,
	theme: ThemePalette,
): SparklineView {
	const colors = resolveColors(group, theme);
	const view: SparklineView = {
		type: group.type,
		color: colors.colorSeries,
		lineWeightPt: group.lineWeight,
		lines: [],
		markers: [],
		columns: [],
	};
	const values = sparklineValues(group, data.values);
	const n = values.length;
	if (!n || !scale) return view;
	const facts = pointFacts(values);
	const xs = positions(group, data, n);
	const flip = (x: number) => (group.rightToLeft ? 1 - x : x);
	const range = scale.max - scale.min;
	const yOf = (v: number) => (range ? Math.max(0, Math.min(1, (scale.max - v) / range)) : 0.5);

	if (group.type === 'line') {
		let segment: SparklinePointView[] = [];
		values.forEach((v, i) => {
			const x = xs[i];
			if (v === null || x === null || x === undefined) {
				if (group.displayEmptyCellsAs === 'gap' && segment.length) {
					view.lines.push(segment);
					segment = [];
				}
				return;
			}
			const point = { x: flip(x), y: yOf(v) };
			segment.push(point);
			const color = pointColor(group, v, i, facts, colors, true);
			if (color) view.markers.push({ ...point, color });
		});
		if (segment.length) view.lines.push(segment);
	} else {
		const slot = 1 / n;
		const w = slot * (1 - COLUMN_GAP);
		// Columns grow from zero, or from the scale's edge when zero is outside it.
		const base = yOf(Math.max(scale.min, Math.min(scale.max, 0)));
		values.forEach((v, i) => {
			const position = xs[i];
			if (v === null || position === null || position === undefined) return;
			// A zero has no height in either column type.
			if (v === 0) return;
			let top: number;
			let bottom: number;
			if (group.type === 'stacked') [top, bottom] = v > 0 ? [0, 0.5] : [0.5, 1];
			// A flat scale: every column fills the cell.
			else if (!range) [top, bottom] = [0, 1];
			else [top, bottom] = [Math.min(base, yOf(v)), Math.max(base, yOf(v))];
			const centre = flip(slot / 2 + position * (1 - slot));
			const color = pointColor(group, v, i, facts, colors, false) ?? view.color;
			view.columns.push({ x: centre - w / 2, y: top, w, h: bottom - top, color });
		});
	}
	if (group.type === 'stacked') {
		if (group.displayXAxis) view.axis = { y: 0.5, color: colors.colorAxis };
	} else if (group.displayXAxis && range && scale.min <= 0 && scale.max >= 0)
		view.axis = { y: yOf(0), color: colors.colorAxis };
	return view;
}
