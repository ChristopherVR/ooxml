/**
 * Horizontal-first bend geometry helpers for multi-segment connectors.
 * OOXML preset paths keep x as the primary axis regardless of aspect ratio;
 * transform flips reverse the endpoints when the shape is mirrored.
 */

/** Segment counts implied by the `bentConnector*` / `curvedConnector*` preset names. */
export type ElbowSegments = 3 | 4 | 5;

/** One `(x, y)` waypoint. */
export type ElbowPoint = readonly [number, number];

/** One cubic-Bezier segment: a single control point (used twice) and an end point. */
export interface ElbowCurveSegment {
	control: ElbowPoint;
	end: ElbowPoint;
}

interface OrientedAxes {
	/** Maps a `(primary, secondary)` pair onto `(x, y)`. */
	toXY: (primary: number, secondary: number) => ElbowPoint;
	primaryStart: number;
	/** The endpoint's coordinate along the primary axis (flip-aware: `startX`/`startY` etc). */
	primaryEnd: number;
	secondaryStart: number;
	secondaryEnd: number;
}

/**
 * OOXML bent and curved connector presets are authored horizontal-first.
 * Their orientation is controlled by transform flips, not box aspect ratio.
 */
export function isHorizontalPrimaryAxis(_width: number, _height: number): boolean {
	return true;
}

function orientAxes(
	startX: number,
	startY: number,
	endX: number,
	endY: number,
	_width: number,
	_height: number,
): OrientedAxes {
	return {
		toXY: (primary, secondary) => [primary, secondary],
		primaryStart: startX,
		primaryEnd: endX,
		secondaryStart: startY,
		secondaryEnd: endY,
	};
}

/**
 * Compute the bend waypoints (including the two endpoints) for a
 * `segments`-segment orthogonal elbow between `(startX,startY)` and
 * `(endX,endY)`, honouring `adj1`/`adj2`/`adj3` fractions (already
 * normalised to 0..1 by `getConnectorAdjustment`).
 *
 * Segment counts mirror the OOXML presets:
 * - `3` (`bentConnector3`, Z-shape): one bend line, positioned by `adj1`.
 * - `4` (`bentConnector4`): a staircase through `adj1` (primary axis) and
 *   `adj2` (secondary axis).
 * - `5` (`bentConnector5`): a staircase with two primary-axis bend lines
 *   (`adj1`, `adj3`) joined by one secondary-axis crossing (`adj2`).
 */
export function elbowWaypoints(
	startX: number,
	startY: number,
	endX: number,
	endY: number,
	width: number,
	height: number,
	segments: ElbowSegments,
	adj1: number,
	adj2: number,
	adj3: number,
): ElbowPoint[] {
	const axes = orientAxes(startX, startY, endX, endY, width, height);
	const { toXY, primaryStart, secondaryStart, secondaryEnd, primaryEnd } = axes;
	const mid1 = primaryStart + (primaryEnd - primaryStart) * adj1;

	if (segments === 3) {
		return [[startX, startY], toXY(mid1, secondaryStart), toXY(mid1, secondaryEnd), [endX, endY]];
	}

	const secMid = secondaryStart + (secondaryEnd - secondaryStart) * adj2;
	if (segments === 4) {
		return [
			[startX, startY],
			toXY(mid1, secondaryStart),
			toXY(mid1, secMid),
			toXY(primaryEnd, secMid),
			[endX, endY],
		];
	}

	const mid2 = primaryStart + (primaryEnd - primaryStart) * adj3;
	return [
		[startX, startY],
		toXY(mid1, secondaryStart),
		toXY(mid1, secMid),
		toXY(mid2, secMid),
		toXY(mid2, secondaryEnd),
		[endX, endY],
	];
}

/**
 * Compute the cubic-Bezier segments (control point used twice, plus an end
 * point) for a `segments`-segment smooth elbow, mirroring
 * {@link elbowWaypoints}'s fixed axis and adjustment handling. `segments`
 * yields `segments - 1` curve segments.
 */
export function elbowCurveSegments(
	startX: number,
	startY: number,
	endX: number,
	endY: number,
	width: number,
	height: number,
	segments: ElbowSegments,
	adj1: number,
	adj2: number,
	adj3: number,
): ElbowCurveSegment[] {
	const axes = orientAxes(startX, startY, endX, endY, width, height);
	const { toXY, primaryStart, secondaryStart, secondaryEnd, primaryEnd } = axes;
	const mid1 = primaryStart + (primaryEnd - primaryStart) * adj1;

	if (segments === 3) {
		const secMid = (secondaryStart + secondaryEnd) / 2;
		return [
			{ control: toXY(mid1, secondaryStart), end: toXY(mid1, secMid) },
			{ control: toXY(mid1, secondaryEnd), end: toXY(primaryEnd, secondaryEnd) },
		];
	}

	const secMid = secondaryStart + (secondaryEnd - secondaryStart) * adj2;
	const quarterSec = secondaryStart + (secMid - secondaryStart) * 0.5;

	if (segments === 4) {
		const midPrimaryBetween = (mid1 + primaryEnd) / 2;
		return [
			{ control: toXY(mid1, secondaryStart), end: toXY(mid1, quarterSec) },
			{ control: toXY(mid1, secMid), end: toXY(midPrimaryBetween, secMid) },
			{ control: toXY(primaryEnd, secMid), end: toXY(primaryEnd, secondaryEnd) },
		];
	}

	const mid2 = primaryStart + (primaryEnd - primaryStart) * adj3;
	const midPrimaryBetween = (mid1 + mid2) / 2;
	const threeQuarterSec = secMid + (secondaryEnd - secMid) * 0.5;
	return [
		{ control: toXY(mid1, secondaryStart), end: toXY(mid1, quarterSec) },
		{ control: toXY(mid1, secMid), end: toXY(midPrimaryBetween, secMid) },
		{ control: toXY(mid2, secMid), end: toXY(mid2, threeQuarterSec) },
		{ control: toXY(mid2, secondaryEnd), end: toXY(primaryEnd, secondaryEnd) },
	];
}

/**
 * Segment count implied by a lower-cased `bentConnector*` / `curvedConnector*`
 * shape type. Unknown or missing suffixes fall back to `3` (the Z-shape).
 */
export function elbowSegmentCount(lowerShapeType: string): ElbowSegments {
	if (lowerShapeType.includes('connector4')) {
		return 4;
	}
	if (lowerShapeType.includes('connector5')) {
		return 5;
	}
	return 3;
}

function fmtPoint(p: ElbowPoint): string {
	return `${p[0]},${p[1]}`;
}

function cubicTo(control: ElbowPoint, end: ElbowPoint): string {
	return `C${fmtPoint(control)} ${fmtPoint(control)} ${fmtPoint(end)}`;
}

/**
 * Render a `segments`-segment elbow as a smooth SVG path: cubic Beziers whose
 * control points sit on the elbow's own corners, so curved connectors get the
 * same horizontal-first routing as {@link elbowWaypoints} without a sharp
 * corner. `3` emits one cubic through the two corner points; `4` and `5`
 * insert one breakpoint per interior corner (halfway along the secondary
 * axis) so the curve bends near the corner instead of overshooting it.
 */
export function curvedElbowPathD(
	startX: number,
	startY: number,
	endX: number,
	endY: number,
	segments: ElbowSegments,
	adj1: number,
	adj2: number,
	adj3: number,
): string {
	const start: ElbowPoint = [startX, startY];
	const mu1 = startX + (endX - startX) * adj1;

	if (segments === 3) {
		return `M${fmtPoint(start)} C${fmtPoint([mu1, startY])} ${fmtPoint([mu1, endY])} ${fmtPoint([endX, endY])}`;
	}

	const mv = startY + (endY - startY) * adj2;
	const vq1 = startY + (mv - startY) * 0.5;

	if (segments === 4) {
		return [
			`M${fmtPoint(start)}`,
			cubicTo([mu1, startY], [mu1, vq1]),
			cubicTo([mu1, mv], [(mu1 + endX) / 2, mv]),
			cubicTo([endX, mv], [endX, endY]),
		].join(' ');
	}

	const mu2 = startX + (endX - startX) * adj3;
	const vq2 = mv + (endY - mv) * 0.5;
	return [
		`M${fmtPoint(start)}`,
		cubicTo([mu1, startY], [mu1, vq1]),
		cubicTo([mu1, mv], [(mu1 + mu2) / 2, mv]),
		cubicTo([mu2, mv], [mu2, vq2]),
		cubicTo([mu2, endY], [endX, endY]),
	].join(' ');
}
