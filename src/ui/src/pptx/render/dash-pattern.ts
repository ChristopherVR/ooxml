/**
 * PowerPoint's dash patterns, measured from its own 1280x720 exports.
 *
 * Every preset is a fixed run of (dash, gap) lengths in units of the pen width:
 *
 * | preset                      | dash, gap, ...                     |
 * | --------------------------- | ---------------------------------- |
 * | `dash`                      | 4 3                                |
 * | `dot`                       | 1 3                                |
 * | `lgDash`                    | 8 3                                |
 * | `dashDot`                   | 4 3 1 3                            |
 * | `lgDashDot`                 | 8 3 1 3                            |
 * | `lgDashDotDot`              | 8 3 1 3 1 3                        |
 * | `sysDash`                   | 3 1                                |
 * | `sysDot`                    | 1 1                                |
 * | `sysDashDot`                | 3 1 1 1                            |
 * | `sysDashDotDot`             | 3 1 1 1 1 1                        |
 * | `custDash` (`a:ds d sp`)    | d, sp in 1000ths of a percent      |
 *
 * Three behaviours the numbers alone do not say:
 *
 *  - The unit is the pen width rounded to WHOLE device pixels (at least one),
 *    the same rounding the line itself gets: at 1280 px a 2 pt line (2.67 px)
 *    dashes as a 3 px pen, 1 pt (1.33 px) as 1 px. See {@link dashPenWidth}.
 *  - The pattern starts at the first point of the path with a dash.
 *  - Caps are part of the dash, not added to it. A `rnd` dash is the pattern's
 *    dash length INCLUDING its round ends (a `dot` is a circle one pen wide and
 *    the period is unchanged), and `sq` adds nothing to interior dashes: they
 *    are the flat dashes, only the line's own ends overshoot. SVG draws a cap
 *    on top of the dash, so {@link capCompensatedDashArray} shortens round dashes
 *    by one pen width, and the stroke uses a `butt` cap for dashed square lines
 *    (`svgLineCap(cap, true)`).
 *
 * Pure and framework-agnostic.
 *
 * @module render/dash-pattern
 */

/** Dash and gap lengths, alternating, in units of the pen width. */
const PRESET_PATTERNS: Readonly<Record<string, readonly number[]>> = {
	dash: [4, 3],
	dot: [1, 3],
	lgDash: [8, 3],
	dashDot: [4, 3, 1, 3],
	lgDashDot: [8, 3, 1, 3],
	lgDashDotDot: [8, 3, 1, 3, 1, 3],
	sysDash: [3, 1],
	sysDot: [1, 1],
	sysDashDot: [3, 1, 1, 1],
	sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** A `custDash` segment: lengths in 1000ths of a percent of the line width. */
export interface DashSegmentInput {
	dash: number;
	space: number;
}

/** `a:ds` lengths are 1000ths of a percent: 100000 is one pen width. */
const CUSTOM_DASH_UNIT = 100000;

/** Used when a `custDash` carries no usable segment (PowerPoint's own default). */
const FALLBACK_CUSTOM_PATTERN: readonly number[] = [3, 2];

/**
 * The pen width dash lengths scale with: the stroke width rounded to whole
 * pixels, at least one (ties round down: 4.5 px dashes as 4 px, measured).
 */
export function dashPenWidth(strokeWidth: number): number {
	if (!Number.isFinite(strokeWidth) || strokeWidth <= 1) return 1;
	return Math.max(1, Math.ceil(strokeWidth - 0.5));
}

/**
 * The dash pattern of a preset or custom dash in pen widths, or `undefined` for
 * a solid line.
 */
export function dashPatternUnits(
	dashType: string | undefined,
	customSegments?: ReadonlyArray<DashSegmentInput>,
): readonly number[] | undefined {
	if (!dashType || dashType === 'solid') return undefined;
	if (dashType === 'custom') {
		const pattern = (customSegments ?? []).flatMap((seg) => [
			Math.max(0, seg.dash) / CUSTOM_DASH_UNIT,
			Math.max(0, seg.space) / CUSTOM_DASH_UNIT,
		]);
		return pattern.some((n) => n > 0) ? pattern : FALLBACK_CUSTOM_PATTERN;
	}
	return PRESET_PATTERNS[dashType];
}

/** Whether an `a:ln` dash token paints a dashed (not solid) line. */
export function isDashedStroke(dashType: string | undefined): boolean {
	return dashType !== undefined && dashType !== 'solid' && dashType !== '';
}

/**
 * The SVG dash array (user units) for a pattern at a stroke width. With a
 * `rnd` cap each dash gives up one pen width to its two round ends, and the gap
 * gains it back, so dots and dashes land where PowerPoint puts them.
 */
export function capCompensatedDashArray(
	pattern: readonly number[],
	strokeWidth: number,
	lineCap?: 'flat' | 'rnd' | 'sq',
): string {
	const pen = dashPenWidth(strokeWidth);
	const round = lineCap === 'rnd';
	const out: number[] = [];
	for (let i = 0; i < pattern.length; i += 1) {
		const len = (pattern[i] ?? 0) * pen;
		if (!round) {
			out.push(len);
		} else if (i % 2 === 0) {
			out.push(Math.max(0, len - pen));
		} else {
			// The shortening of the dash before it is added here.
			const dash = (pattern[i - 1] ?? 0) * pen;
			out.push(len + (dash - Math.max(0, dash - pen)));
		}
	}
	return out.map((n) => Number(n.toFixed(4))).join(' ');
}
