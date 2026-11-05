/**
 * The turn direction of a morph: the short way round, as PowerPoint takes it.
 *
 * @module render/morph-rotation
 */

/**
 * The target angle to animate TO so the element turns the short way round.
 *
 * CSS interpolates `rotate(a)` -> `rotate(b)` numerically, so a pair authored
 * at 315deg and 0deg spins -315deg (almost a full turn anti-clockwise) when
 * the shapes are only 45deg apart. PowerPoint always takes the shorter arc.
 * Returns `fromDeg` plus the delta wrapped into (-180, 180], which is the same
 * final orientation modulo 360 but the rotation a viewer expects: the issue
 * #131 deck's wheel points its arrow at the selected wedge by rotating a ring
 * in 45deg steps, so clicking the neighbouring wedge sent the arrow the long
 * way around the dial.
 *
 * An exact half turn has no shorter arc, and PowerPoint's choice there is not
 * a fixed sign: it turns CLOCKWISE when the shape starts anywhere in
 * [90, 270) and ANTI-clockwise otherwise. Measured on PowerPoint 16 by
 * sampling the rendered frames of a half-turn morph, both on the issue #131
 * wheel (0->180 anti, 45->225 anti, 90->270 clock, 135->315 clock, 180->360
 * clock, 270->90 anti) and on a synthetic two-slide deck built for the purpose
 * (0->180 anti, 45->225 anti, 90->270 clock, 270->90 anti), which agrees on
 * every case. Always taking +180 sent the wheel's arrow round the wrong side
 * for the wedge diametrically opposite the one on screen, so one click in
 * seven looked nothing like the others.
 */
export function shortestRotationTarget(fromDeg: number, toDeg: number): number {
	let delta = (toDeg - fromDeg) % 360;
	if (delta > 180) {
		delta -= 360;
	} else if (delta < -180) {
		delta += 360;
	}
	if (Math.abs(Math.abs(delta) - 180) < 1e-9) {
		const start = ((fromDeg % 360) + 360) % 360;
		delta = start >= 90 && start < 270 ? 180 : -180;
	}
	return fromDeg + delta;
}
