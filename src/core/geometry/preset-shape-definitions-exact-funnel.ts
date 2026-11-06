/**
 * Geometry of `funnel`: a rim ellipse `wd2` x `hd4`, an opening inset `ss / 20`
 * (left empty), and a small bottom ellipse joined to the rim by two straight
 * sides that are tangent to both ellipses (measured on PowerPoint 16 at five
 * sizes; the shape has no adjustments).
 *
 * The common tangent of two axis-aligned ellipses has no closed form, so the
 * side slope `m` (the run per unit of height) is found by six Newton steps on
 * `m * (cy2 - cy1) = sqrt(a1^2 + m^2 b1^2) - sqrt(a2^2 + m^2 b2^2)`.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

/** One Newton step from slope guide `m` to guide `n`. */
function newton(m: string, n: string): string[] {
	return [
		`${n}b1 */ ${m} b1 1`,
		`${n}b2 */ ${m} b2 1`,
		`${n}s1 mod a1 ${n}b1 0`,
		`${n}s2 mod a2 ${n}b2 0`,
		`${n}m1 */ ${m} dY 1`,
		`${n}f0 +- ${n}m1 ${n}s2 ${n}s1`,
		`${n}t1 */ ${n}b1 b1 ${n}s1`,
		`${n}t2 */ ${n}b2 b2 ${n}s2`,
		`${n}d0 +- dY ${n}t2 ${n}t1`,
		`${n}d1 */ ${n}f0 1 ${n}d0`,
		`${n} +- ${m} 0 ${n}d1`,
	];
}

const STEPS = ['mA', 'mB', 'mC', 'mD', 'mE', 'mF'];

const funnel = exactDef(
	'funnel',
	undefined,
	[
		'a1 */ wd2 1 1',
		'b1 */ hd4 1 1',
		'a2 */ w 1 9',
		'b2 */ h 1 23',
		'cy2 +- b 0 b2',
		'dY +- cy2 0 hd4',
		'a1m +- a1 0 a2',
		'm0 */ a1m 1 dY',
		...newton('m0', 'mA'),
		...STEPS.slice(1).flatMap((n, i) => newton(STEPS[i]!, n)),
		'mb1 */ mF b1 1',
		'mb2 */ mF b2 1',
		'phi1 at2 a1 mb1',
		'phi2 at2 a2 mb2',
		// left tangent point on the rim, offsets from the centre
		'r1x cos a1 phi1',
		'r1y sin b1 phi1',
		'r1nx +- 0 0 r1x',
		'x1l +- hc 0 r1x',
		'y1l +- hd4 r1y 0',
		'x1r +- hc r1x 0',
		'ang1 at2 r1nx r1y',
		'ang1x2 */ ang1 2 1',
		'sw1 +- 32400000 0 ang1x2',
		'ang1r +- 10800000 0 ang1',
		// right tangent point on the bottom ellipse
		'r2x cos a2 phi2',
		'r2y sin b2 phi2',
		'x2r +- hc r2x 0',
		'y2r +- cy2 r2y 0',
		'ang2r at2 r2x r2y',
		'ang2x2 */ ang2r 2 1',
		'sw2 +- 10800000 0 ang2x2',
		// opening
		'dd */ ss 1 20',
		'iw +- wd2 0 dd',
		'ih +- hd4 0 dd',
		'ix +- hc iw 0',
	],
	'l t r b',
	[p('M x1l y1l A a1 b1 ang1 sw1 L x2r y2r A a2 b2 ang2r sw2 Z M ix hd4 A iw ih 0 -21600000 Z')],
);

export const EXACT_FUNNEL_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	funnel,
};
