import { visioOpenArrowExtent } from './open-arrow';

/** Native measured glyph with an endpoint setback, in local y-up inches. */
export interface VisioFilledArrow {
	path: string;
	setback: number;
	beginSetback: number;
	/** True for a filled glyph; false for a closed outline drawn with the line's stroke. */
	filled: boolean;
}

export interface VisioArrowLineLayout {
	path: string;
	startSetback: number;
	endSetback: number;
}

/** Native stem and marker anchors for a straight filled connector. */
export function layoutVisioFilledArrowLine(
	path: string,
	start: VisioFilledArrow | undefined,
	end: VisioFilledArrow | undefined,
): VisioArrowLineLayout | undefined {
	for (const arrow of [start, end]) {
		if (
			arrow &&
			(!Number.isFinite(arrow.setback) ||
				!Number.isFinite(arrow.beginSetback) ||
				arrow.setback < 0 ||
				arrow.beginSetback < 0 ||
				arrow.beginSetback > arrow.setback)
		)
			return undefined;
	}
	const coordinates = straightLine(path);
	if (!coordinates || (!start && !end)) return undefined;
	const [x, y, ex, ey] = coordinates;
	const length = Math.hypot(ex - x, ey - y);
	if (!Number.isFinite(length) || length === 0) return undefined;
	if (!start && end) {
		const difference = length - end.setback;
		if (Math.abs(difference) <= 1e-9) {
			const extension = end.setback - end.beginSetback;
			return {
				path: `M ${x} ${y} L ${x} ${y} L ${x + ((ex - x) * extension) / length} ${y + ((ey - y) * extension) / length}`,
				startSetback: 0,
				endSetback: length - extension,
			};
		}
		if (difference < 0) return { path, startSetback: 0, endSetback: 0 };
	}
	const trimmed = trimVisioArrowLine(
		path,
		start?.beginSetback ?? 0,
		end?.setback ?? 0,
		start ? start.setback - start.beginSetback : 0,
	);
	return trimmed
		? { path: trimmed, startSetback: start?.beginSetback ?? 0, endSetback: end?.setback ?? 0 }
		: undefined;
}

function straightLine(path: string): readonly [number, number, number, number] | undefined {
	const n = '([-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][-+]?\\d+)?)';
	const match = new RegExp(`^\\s*M\\s*${n}[ ,]+${n}\\s*L\\s*${n}[ ,]+${n}\\s*$`).exec(path);
	return match ? (match.slice(1).map(Number) as [number, number, number, number]) : undefined;
}
// Native numerical geometry at unit scale, read from Visio 16 SVG exports of BeginArrow/EndArrow
// codes 1 to 45 (glyph, then the setback the export trims the line by, in glyph units). Closed
// fill implicitly closes each path. Codes whose export repeats another code's glyph share it.
const DOT =
	'M 0 0.75 C -0.414214 0.75 -0.75 0.414214 -0.75 0 -0.75 -0.414214 -0.414214 -0.75 0 -0.75 0.414214 -0.75 0.75 -0.414214 0.75 0 0.75 0.414214 0.414214 0.75 0 0.75 Z';
const SQUARE = 'M 0.75 -0.75 L 0.75 0.75 L -0.75 0.75 L -0.75 -0.75 L 0.75 -0.75';
const RING =
	'M 0.75 0.75 C 0.335786 0.75 0 0.414214 0 0 0 -0.414214 0.335786 -0.75 0.75 -0.75 1.16421 -0.75 1.5 -0.414214 1.5 0 1.5 0.414214 1.16421 0.75 0.75 0.75 Z';
const FAR_RING =
	'M 2.25 0.75 C 1.83579 0.75 1.5 0.414214 1.5 0 1.5 -0.414214 1.83579 -0.75 2.25 -0.75 2.66421 -0.75 3 -0.414214 3 0 3 0.414214 2.66421 0.75 2.25 0.75 Z';
const CONCAVE =
	'M 2 1 L 0 0 L 1.98117 -0.993387 C 1.67173 -0.364515 1.67301 0.372641 1.98465 1.00043';
const CONVEX = 'M 2 1 L 0 0 L 2.01237 -1.00641 C 2.32921 -0.372876 2.32921 0.372869 2.01238 1.0064';
const DOUBLE = 'M 0 0 L 2 -1 L 2 1 L 0 0 M 2 0 L 4 -1 L 4 1 L 2 0';
const BARS = (count: number) =>
	Array.from(
		{ length: count },
		(_, i) => `M ${2.25 + i * 0.75} -0.75 L ${2.25 + i * 0.75} 0.75`,
	).join(' ');
const RING_DIAMOND = `${RING} M 1.5 0 L 3.75 -1.15 L 6 0 L 3.75 1.15 L 1.5 0`;
const glyphs: Readonly<Record<number, readonly [string, number]>> = {
	2: ['M 1 1 L 0 0 L 1 -1 L 1 1', 1],
	4: ['M 2 1 L 0 0 L 2 -1 L 2 1', 2],
	5: [CONCAVE, 1.75],
	6: [CONVEX, 2.25],
	8: [CONCAVE, 1.75],
	10: [DOT, 0.75],
	11: [SQUARE, 0.75],
	13: ['M 3 1 L 0 0 L 3 -1 L 3 1', 3],
	39: [DOUBLE, 4],
	42: [DOT, 0.75],
};
const outlines: Readonly<Record<number, readonly [string, number]>> = {
	14: ['M 3 -1 L 0 0 L 3 1 L 3 -1', 3],
	15: ['M 1 -1 L 0 0 L 1 1 L 1 -1', 1],
	16: ['M 0 0 L 2 -1 L 2 1 L 0 0', 2],
	17: [
		'M 2 1 L 0 0 L 1.98117 -0.993387 C 1.67051 -0.362037 1.67309 0.378293 1.98815 1.00746',
		1.75,
	],
	18: [CONVEX, 2.25],
	19: [
		'M 1.95144 1.125 C 1.55653 0.440984 1.54919 -0.400016 1.93211 -1.09082 L 1.92863 -1.09117 C 1.52903 -0.426131 0.815005 -0.0138845 0.0392708 -0.000343084 L 0.0392666 0.000342369 C 0.815001 0.0138824 1.52903 0.426127 1.92863 1.09116',
		1.65,
	],
	20: [DOT, 0.75],
	21: [SQUARE, 0.75],
	22: ['M 0 0 L 2.25 -1.125 L 4.5 0 L 2.25 1.125 L 0 0', 4.5],
	29: [`${FAR_RING} M 0 -0.75 L 1.5 0 L 0 0.75 M 1.5 0 L 0 0`, 3],
	30: [`${FAR_RING} M 0.75 -0.75 L 0.75 0.75 M 1.5 0 L 0 0`, 3],
	31: [`${RING} ${BARS(1)}`, 1.5],
	32: [`${RING} ${BARS(2)}`, 1.5],
	33: [`${RING} ${BARS(3)}`, 1.5],
	34: [RING_DIAMOND, 6],
	35: [`${RING} ${BARS(1)}`, 1.5],
	36: [`${RING} ${BARS(2)}`, 1.5],
	37: [`${RING} ${BARS(3)}`, 1.5],
	38: [RING_DIAMOND, 6],
	40: [DOUBLE, 4],
	41: [RING, 1.5],
};

/** True when a BeginArrow/EndArrow code has a native glyph with a setback (filled or outline). */
export function visioHasSetbackArrow(code: number): boolean {
	return code in glyphs || code in outlines;
}

export function visioFilledArrow(
	code: number,
	size: number,
	lineWidth: number,
): VisioFilledArrow | undefined {
	const filled = glyphs[code];
	const glyph = filled ?? outlines[code];
	if (!glyph) return undefined;
	const extent = visioOpenArrowExtent(size, lineWidth);
	const setback = glyph[1] * extent;
	let index = 0;
	return {
		path: glyph[0].replace(/-?\d+(?:\.\d+)?/g, (n) =>
			String(Number(n) * extent * (index++ % 2 ? 1 : -1)),
		),
		setback,
		beginSetback: setback - Math.max(0.005, setback * 0.01),
		filled: !!filled,
	};
}

/** Trim a single straight open line for native filled markers. Reject other paths
 * and overlapping setbacks instead of silently approximating their tangents.
 */
export function trimVisioArrowLine(
	path: string,
	start: number,
	end: number,
	beginExtension = 0,
): string | undefined {
	const coordinates = straightLine(path);
	if (!coordinates || ![start, end, beginExtension].every((v) => Number.isFinite(v) && v >= 0))
		return undefined;
	const [x, y, ex, ey] = coordinates;
	const dx = ex - x,
		dy = ey - y,
		length = Math.hypot(dx, dy);
	const base = start > 0 ? start + beginExtension : 0;
	if (!Number.isFinite(length) || length === 0 || base + end >= length) return undefined;
	const point = (distance: number) =>
		`${x + (dx * distance) / length} ${y + (dy * distance) / length}`;
	return `M ${point(start)}${base > start ? ` L ${point(base)}` : ''} L ${point(length - end)}`;
}
