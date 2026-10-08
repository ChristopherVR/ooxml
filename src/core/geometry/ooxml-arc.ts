/**
 * OOXML `a:arcTo` to SVG arc conversion, shared by the preset shape evaluator
 * and custom geometry. Moved from `pptx/core/geometry/guide-formula-paths.ts`.
 */

import { angleToRadians } from './guide-formula-eval';

/**
 * Result of converting an OOXML `a:arcTo` command to SVG arc notation.
 */
export interface ArcToResult {
	/** SVG arc path segment string (e.g. `"A 50 30 0 1 1 100 75"`). */
	svg: string;
	/** X coordinate of the arc endpoint. */
	endX: number;
	/** Y coordinate of the arc endpoint. */
	endY: number;
}

/**
 * Convert an OOXML `a:arcTo` command to an SVG arc path segment.
 *
 * OOXML arcTo: the current pen position lies on an implicit ellipse at
 * angle `stAng`. The arc sweeps `swAng` degrees (in 60000ths). The
 * implicit ellipse center is derived from the current position and stAng.
 *
 * @param wR - Horizontal radius of the ellipse.
 * @param hR - Vertical radius of the ellipse.
 * @param stAng - Start angle in 60000ths of a degree.
 * @param swAng - Sweep angle in 60000ths of a degree.
 * @param penX - Current pen X position.
 * @param penY - Current pen Y position.
 */
export function ooxmlArcToSvg(
	wR: number,
	hR: number,
	stAng: number,
	swAng: number,
	penX: number,
	penY: number,
): ArcToResult | null {
	// Degenerate arcs: zero radius or zero sweep produce no visible arc
	if (wR <= 0 || hR <= 0 || swAng === 0) {
		return null;
	}

	// Convert OOXML angles (60,000ths of a degree) to radians
	const startRad = angleToRadians(stAng);
	const sweepRad = angleToRadians(swAng);
	const endRad = startRad + sweepRad;

	// Derive the implicit ellipse center from the current pen position.
	// The pen sits on the ellipse at the start angle, so:
	//   penX = cx + wR * cos(startRad)  =>  cx = penX - wR * cos(startRad)
	//   penY = cy + hR * sin(startRad)  =>  cy = penY - hR * sin(startRad)
	const cx = penX - wR * Math.cos(startRad);
	const cy = penY - hR * Math.sin(startRad);

	// Compute the absolute endpoint on the ellipse at the end angle
	const endX = cx + wR * Math.cos(endRad);
	const endY = cy + hR * Math.sin(endRad);

	// SVG arc flags:
	// - large-arc-flag: 1 if the arc spans more than 180 degrees
	// - sweep-flag: 1 if the arc is drawn in the positive-angle direction
	const largeArc = Math.abs(sweepRad) > Math.PI ? 1 : 0;
	const sweep = sweepRad > 0 ? 1 : 0;

	// Round to 3 decimal places for clean SVG output
	const rx = Math.round(wR * 1000) / 1000;
	const ry = Math.round(hR * 1000) / 1000;
	const ex = Math.round(endX * 1000) / 1000;
	const ey = Math.round(endY * 1000) / 1000;

	return {
		svg: `A ${rx} ${ry} 0 ${largeArc} ${sweep} ${ex} ${ey}`,
		endX,
		endY,
	};
}
