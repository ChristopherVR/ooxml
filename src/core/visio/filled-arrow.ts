import { visioOpenArrowExtent } from './open-arrow.js';

/** Native measured filled glyph and endpoint setback, in local y-up inches. */
export interface VisioFilledArrow {
	path: string;
	setback: number;
	beginSetback: number;
}
// Native numerical geometry at unit scale. Closed fill implicitly closes each path.
const glyphs: Readonly<Record<number, readonly [string, number]>> = {
	2: ['M 1 1 L 0 0 L 1 -1 L 1 1', 1],
	4: ['M 2 1 L 0 0 L 2 -1 L 2 1', 2],
	5: ['M 2 1 L 0 0 L 1.98117 -0.993387 C 1.67173 -0.364515 1.67301 0.372641 1.98465 1.00043', 1.75],
	6: ['M 2 1 L 0 0 L 2.01237 -1.00641 C 2.32921 -0.372876 2.32921 0.372869 2.01238 1.0064', 2.25],
};

export function visioFilledArrow(
	code: number,
	size: number,
	lineWidth: number,
): VisioFilledArrow | undefined {
	const glyph = glyphs[code];
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
	const n = '([-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][-+]?\\d+)?)';
	const match = new RegExp(`^\\s*M\\s*${n}[ ,]+${n}\\s*L\\s*${n}[ ,]+${n}\\s*$`).exec(path);
	if (!match || ![start, end, beginExtension].every((v) => Number.isFinite(v) && v >= 0))
		return undefined;
	const [x, y, ex, ey] = match.slice(1).map(Number) as [number, number, number, number];
	const dx = ex - x,
		dy = ey - y,
		length = Math.hypot(dx, dy);
	const base = start > 0 ? start + beginExtension : 0;
	if (!Number.isFinite(length) || length === 0 || base + end >= length) return undefined;
	const point = (distance: number) =>
		`${x + (dx * distance) / length} ${y + (dy * distance) / length}`;
	return `M ${point(start)}${base > start ? ` L ${point(base)}` : ''} L ${point(length - end)}`;
}
