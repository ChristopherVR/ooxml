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

/** Parametric angle of the ellipse point seen at visual angle `visual` (radians). */
export function visualToParametric(visual: number, wR: number, hR: number): number {
	return Math.atan2(wR * Math.sin(visual), hR * Math.cos(visual));
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
 * @param scaleX - Horizontal scale from the radii's own space (a path's `w`
 *   space) to the output. The visual-to-parametric angle conversion has to run
 *   in the radii's own space (where a `cloud` lobe is a circle), then scale.
 * @param scaleY - Vertical counterpart of `scaleX`.
 * @param parametric - Treat `stAng` / `swAng` as parametric angles instead of
 *   visual ones. Only for approximate hand-written definitions that were tuned
 *   that way and have no spec-exact replacement yet.
 */
export function ooxmlArcToSvg(
	wR: number,
	hR: number,
	stAng: number,
	swAng: number,
	penX: number,
	penY: number,
	scaleX = 1,
	scaleY = 1,
	parametric = false,
): ArcToResult | null {
	// Degenerate arcs: zero radius or zero sweep produce no visible arc
	if (
		![wR, hR, stAng, swAng, penX, penY].every(Number.isFinite) ||
		wR <= 0 ||
		hR <= 0 ||
		swAng === 0
	) {
		return null;
	}
	// SVG omits an arc whose endpoints coincide. A full DrawingML turn needs two arcs.
	if (Math.abs(swAng) >= 21600000) {
		const sweep = Math.sign(swAng) * (21600000 + (Math.abs(swAng) % 21600000));
		const pieces = Math.ceil(Math.abs(sweep) / 10800000);
		let x = penX;
		let y = penY;
		let angle = stAng;
		const commands: string[] = [];
		for (let index = 0; index < pieces; index++) {
			const part = ooxmlArcToSvg(wR, hR, angle, sweep / pieces, x, y, scaleX, scaleY, parametric)!;
			commands.push(part.svg);
			x = part.endX;
			y = part.endY;
			angle += sweep / pieces;
		}
		return { svg: commands.join(' '), endX: x, endY: y };
	}

	// ECMA-376 arc angles are VISUAL angles: the direction of the arc point as
	// seen from the ellipse centre. On a circle that equals the parametric angle
	// `t` of `(cx + wR cos t, cy + hR sin t)`; on an ellipse (every non-square
	// preset) `tan t = (wR / hR) tan a`. PowerPoint places arc end points this
	// way (measured against `gear6`, `curvedRightArrow`, `circularArrow`), so
	// convert both ends and keep the sweep in the direction that was asked for.
	const visualStart = angleToRadians(stAng);
	const visualSweep = angleToRadians(swAng);
	let startRad = visualStart;
	let endRad = visualStart + visualSweep;
	if (!parametric) {
		startRad = visualToParametric(visualStart, wR, hR);
		endRad = visualToParametric(visualStart + visualSweep, wR, hR);
		const direction = Math.sign(swAng);
		while ((endRad - startRad) * direction < 0) {
			endRad += direction * 2 * Math.PI;
		}
		while (Math.abs(endRad - startRad) > Math.abs(visualSweep) + Math.PI) {
			endRad -= direction * 2 * Math.PI;
		}
	}
	const sweepRad = endRad - startRad;

	// Derive the implicit ellipse center from the current pen position.
	// The pen sits on the ellipse at the start angle, so:
	//   penX = cx + wR * cos(startRad)  =>  cx = penX - wR * cos(startRad)
	//   penY = cy + hR * sin(startRad)  =>  cy = penY - hR * sin(startRad)
	const outWR = wR * scaleX;
	const outHR = hR * scaleY;
	const cx = penX - outWR * Math.cos(startRad);
	const cy = penY - outHR * Math.sin(startRad);

	// Compute the absolute endpoint on the ellipse at the end angle
	const endX = cx + outWR * Math.cos(endRad);
	const endY = cy + outHR * Math.sin(endRad);

	// SVG arc flags:
	// - large-arc-flag: 1 if the arc spans more than 180 degrees
	// - sweep-flag: 1 if the arc is drawn in the positive-angle direction
	const largeArc = Math.abs(sweepRad) > Math.PI ? 1 : 0;
	const sweep = sweepRad > 0 ? 1 : 0;

	// Round to 3 decimal places for clean SVG output
	const rx = Math.round(outWR * 1000) / 1000;
	const ry = Math.round(outHR * 1000) / 1000;
	const ex = Math.round(endX * 1000) / 1000;
	const ey = Math.round(endY * 1000) / 1000;

	return {
		svg: `A ${rx} ${ry} 0 ${largeArc} ${sweep} ${ex} ${ey}`,
		endX,
		endY,
	};
}
