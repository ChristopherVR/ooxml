/**
 * Dash patterns for connector lines.
 *
 * Split out of `connector-path.ts` alongside the arrow markers: a connector's
 * stroke decoration is a separate concern from where the line is routed, and
 * both are configured independently in the inspector.
 *
 * Pure and framework-agnostic.
 */

import { squareCapDashedPathData } from './dash-square-cap';
import { getSvgStrokeDasharray, normalizeStrokeDashType } from './element-style-transform';

/** A single custom-dash segment (percent-of-line-width, 1000ths of a percent). */
export interface DashSegment {
	dash: number;
	space: number;
}

/**
 * Return the SVG `stroke-dasharray` string for a given OOXML stroke dash preset
 * and width, or `undefined` for solid lines (no attribute needed).
 *
 * Produces a distinct pattern per preset (`dash`, `lgDash`, `dashDot`,
 * `sysDashDotDot`, etc.) rather than collapsing every non-dot preset to a single
 * `3w/w` approximation, and honours a `custDash` segment list (`a:custDash/a:ds`)
 * when supplied. This delegates to the same {@link getSvgStrokeDasharray} the
 * shape/border code uses, so connectors and shape outlines stay in lock-step.
 *
 * @param dash               Raw `a:ln/@prstDash` token (e.g. `"lgDashDot"`).
 * @param strokeWidth        Resolved stroke width in px.
 * @param customDashSegments Optional `custDash` segments; when present they take
 *                           precedence and are rendered as an explicit pattern.
 * @param lineCap            `a:ln/@cap`: a round cap is part of each dash.
 */
export function buildDashArray(
	dash: string | undefined,
	strokeWidth: number,
	customDashSegments?: ReadonlyArray<DashSegment>,
	lineCap?: 'flat' | 'rnd' | 'sq',
): string | undefined {
	const segments =
		customDashSegments && customDashSegments.length > 0
			? customDashSegments.map((seg) => ({ dash: seg.dash, space: seg.space }))
			: undefined;
	// A `custDash` implies the `custom` dash family even when no `@prstDash`
	// token was authored alongside it.
	const dashType = segments ? 'custom' : normalizeStrokeDashType(dash);
	return getSvgStrokeDasharray(dashType, strokeWidth, segments, lineCap);
}

/** The `a:ln` fields {@link connectorSquareCapPath} reads. */
export interface ConnectorDashStyle {
	strokeDash?: string | undefined;
	lineCap?: 'flat' | 'rnd' | 'sq' | undefined;
	customDashSegments?: ReadonlyArray<DashSegment> | undefined;
}

/**
 * The path data a connector's painted line must use to show PowerPoint's
 * overshoot at the two ends of a dashed square-cap line (see `dash-square-cap`),
 * or `undefined` when its own geometry is right as it is. `d` is the connector's
 * absolute `M`/`L` path (`M x1,y1 L x2,y2` for a straight one). A line with
 * arrowheads or several strands is left alone: markers would attach to the
 * extra sub-paths.
 */
export function connectorSquareCapPath(
	d: string,
	style: ConnectorDashStyle | undefined,
	strokeWidth: number,
	hasMarkersOrStrands: boolean,
): string | undefined {
	if (hasMarkersOrStrands || style?.lineCap !== 'sq' || !style.strokeDash) return undefined;
	if (normalizeStrokeDashType(style.strokeDash) === 'solid') return undefined;
	const out = squareCapDashedPathData(
		d,
		Math.max(strokeWidth, 1),
		style.strokeDash,
		style.lineCap,
		style.customDashSegments,
	);
	return out === d ? undefined : out;
}
