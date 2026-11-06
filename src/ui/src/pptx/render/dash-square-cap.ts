/**
 * Where PowerPoint puts the ends of a DASHED line with square caps.
 *
 * Measured on its own 1280x720 exports (0.75 pt to 9 pt, all ten presets): a
 * square cap is applied to the line, not to every dash. The interior dashes
 * are the flat dashes of the pattern; the FIRST dash reaches back past the
 * start of the line by half the pen width, and the LAST one reaches past the
 * end by half the pen width when the line ends inside a dash (a line that ends
 * in a gap gains nothing). SVG's own `square` cap would lengthen every dash, so
 * a dashed square line is stroked with `butt` caps (`svgLineCap(cap, true)`)
 * and these two overshoots are drawn as extra sub-paths: a dash pattern
 * restarts at every sub-path, and each sub-path here is shorter than the
 * pattern's first dash, so it is painted solid.
 *
 * Only a single sub-path of absolute `M`/`L` (a straight line or a polyline,
 * which covers straight and bent connectors and the line presets) is
 * rewritten; curves, closed outlines and anything else come back unchanged.
 *
 * Pure and framework-agnostic.
 *
 * @module render/dash-square-cap
 */

import { dashPatternUnits, dashPenWidth } from './dash-pattern';
import type { DashSegmentInput } from './dash-pattern';
import { normalizeStrokeDashType } from './element-style-transform';

type Point = readonly [number, number];

function parsePolyline(d: string): Point[] | undefined {
	const tokens = d.match(/[A-Za-z]|-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/giu);
	if (!tokens || tokens[0] !== 'M') return undefined;
	const points: Point[] = [];
	let cmd = '';
	let i = 0;
	while (i < tokens.length) {
		const token = tokens[i] as string;
		if (/^[A-Za-z]$/u.test(token)) {
			cmd = token;
			i += 1;
			if (cmd !== 'M' && cmd !== 'L') return undefined;
			if (cmd === 'M' && points.length > 0) return undefined;
			continue;
		}
		const x = Number(token);
		const y = Number(tokens[i + 1]);
		if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined;
		points.push([x, y]);
		i += 2;
	}
	return points.length >= 2 ? points : undefined;
}

const fmt = (n: number): string => String(Number(n.toFixed(4)));

/** Whether a line `length` long, dashed from its start with `lengths`, ends inside a dash. */
function endsInDash(lengths: readonly number[], length: number): boolean {
	const period = lengths.reduce((a, b) => a + b, 0);
	if (period <= 0) return false;
	let at = length % period;
	for (let i = 0; i < lengths.length; i += 1) {
		const len = lengths[i] as number;
		if (at < len) return i % 2 === 0 && at > 0;
		at -= len;
	}
	return false;
}

/**
 * The path data of a dashed square-cap line with PowerPoint's overshoot at its
 * two ends (see the module comment), or `d` unchanged when it does not apply.
 *
 * @param d           Absolute `M`/`L` path data of one open sub-path.
 * @param strokeWidth The authored stroke width in px.
 * @param dashType    The `a:prstDash` token (`custom` with `customSegments`).
 * @param lineCap     `a:ln/@cap`; anything but `sq` returns `d`.
 */
export function squareCapDashedPathData(
	d: string,
	strokeWidth: number,
	dashType: string | undefined,
	lineCap: string | undefined,
	customSegments?: ReadonlyArray<DashSegmentInput>,
): string {
	if (lineCap !== 'sq') return d;
	const pattern = dashPatternUnits(
		customSegments && customSegments.length > 0 ? 'custom' : normalizeStrokeDashType(dashType),
		customSegments,
	);
	const points = parsePolyline(d);
	if (!pattern || !points || !(strokeWidth > 0)) return d;
	const pen = dashPenWidth(strokeWidth);
	const half = pen / 2;
	const unit = pattern.map((n) => n * pen);
	// An odd-length dash array is repeated, so the pattern alternates dash and gap.
	const lengths = unit.length % 2 === 1 ? [...unit, ...unit] : unit;
	if ((lengths[0] ?? 0) < half) return d;

	let total = 0;
	let first: Point | undefined;
	let last: Point | undefined;
	for (let i = 1; i < points.length; i += 1) {
		const a = points[i - 1] as Point;
		const b = points[i] as Point;
		const len = Math.hypot(b[0] - a[0], b[1] - a[1]);
		if (len === 0) continue;
		total += len;
		first ??= [(b[0] - a[0]) / len, (b[1] - a[1]) / len];
		last = [(b[0] - a[0]) / len, (b[1] - a[1]) / len];
	}
	if (!first || !last || total === 0) return d;

	const p0 = points[0] as Point;
	const pn = points[points.length - 1] as Point;
	let out = `M${fmt(p0[0] - first[0] * half)},${fmt(p0[1] - first[1] * half)} L${fmt(p0[0])},${fmt(p0[1])} ${d}`;
	if (endsInDash(lengths, total)) {
		out += ` M${fmt(pn[0])},${fmt(pn[1])} L${fmt(pn[0] + last[0] * half)},${fmt(pn[1] + last[1] * half)}`;
	}
	return out;
}
