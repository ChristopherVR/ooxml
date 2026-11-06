/**
 * Geometry of `ellipseRibbon` and `ellipseRibbon2` (the curved banners).
 *
 * Every curved edge is a piece of one parabola `P(x) = f1 * (x - x^2 / w)` that
 * sags `dy1` at the centre, drawn as a quadratic Bezier (a quadratic Bezier with
 * its control point halfway along x is exactly a parabola). The tails hang from
 * the end of the top edge down to `b - dy2`; the centre banner spans from `dy2`
 * at the middle to the bottom edge, and the two folds are the triangles joining
 * a tail's top edge to the banner.
 *
 * `ellipseRibbon2` is `ellipseRibbon` turned over, so both are written once
 * with each y given as an offset from the top that is flipped for the second.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const AV = { adj1: 25000, adj2: 50000, adj3: 12500 };

/** Parabola value at `x` (the guide name `n` holds the result). */
function para(n: string, x: string): string[] {
	return [`${n}q */ ${x} ${x} w`, `${n}r +- ${x} 0 ${n}q`, `${n} */ f1 ${n}r 1`];
}

function build(name: string, flip: boolean): PresetShapeGeometryDefinition {
	/** A y guide from an offset (distance from the reference edge). */
	const y = (n: string, off: string) => (flip ? `${n} +- b 0 ${off}` : `${n} +- t ${off} 0`);
	const gd = [
		'a1 pin 0 adj1 100000',
		'a2 pin 25000 adj2 75000',
		'q10 +- 100000 0 a1',
		'q11 */ q10 1 2',
		'q12 +- a1 0 q11',
		'minAdj3 max 0 q12',
		'a3 pin minAdj3 adj3 a1',
		'dx2 */ w a2 200000',
		'x2 +- hc 0 dx2',
		'x3 +- x2 wd8 0',
		'x9 +- r 0 x2',
		'x5 +- r 0 x3',
		'x10 +- r 0 wd8',
		'dy1 */ h a3 100000',
		'dy2 */ h a1 100000',
		'f1 */ 4 dy1 w',
		'yc +- dy2 0 dy1',
		'yb +- h 0 dy2',
		...para('s3', 'x3'),
		...para('s2', 'x2'),
		...para('sn', 'wd8'),
		'cx1 */ x3 1 2',
		'cx1r +- r 0 cx1',
		'cx2 */ x2 1 2',
		'cx2r +- r 0 cx2',
		'xf +- x2 x3 0',
		'xfm */ xf 1 2',
		'xfmr +- r 0 xfm',
		...para('sf', 'xfm'),
		// offsets from the reference edge
		'oS3 */ s3 1 1',
		'oC3 +- s3 yc 0',
		'oC2 +- s2 yc 0',
		'oB2 +- yb s2 0',
		'oD2 +- h 0 dy1',
		'oD2s +- oD2 s2 0',
		'oN */ yb 1 2',
		'oNn +- oN sn 0',
		'oCy1 */ f1 cx1 1',
		'oCy2b */ f1 cx2 1',
		'oCy2c +- yb oCy2b 0',
		'oCt +- dy2 dy2 oC3',
		'oCb +- h dy1 s2',
		y('yZ', '0'),
		y('yS3', 'oS3'),
		y('yC3', 'oC3'),
		y('yC2', 'oC2'),
		y('yB2', 'oB2'),
		y('yD2', 'oD2s'),
		y('yN', 'oNn'),
		y('yBase', 'yb'),
		y('yCy1', 'oCy1'),
		y('yCy2', 'oCy2c'),
		y('yCt', 'oCt'),
		y('yCb', 'oCb'),
	];
	// Control y of a fold's lower edge, from the exact quadratic Bezier identity
	// C = 2 Pm - (P0 + P1) / 2 for a parabola piece with mid value Pm.
	gd.push(
		'oFe +- oC3 oC2 0',
		'oFe2 */ oFe 1 2',
		'oFmid +- sf yc 0',
		'oFctl */ oFmid 2 1',
		'oFcy +- oFctl 0 oFe2',
		y('yFcy', 'oFcy'),
	);
	return exactDef(name, AV, gd, 'x2 t x5 b', [
		p(
			'M l yZ Q cx1 yCy1 x3 yS3 L x3 yC3 Q hc yCt x5 yC3 L x5 yS3 Q cx1r yCy1 r yZ ' +
				'L x10 yN L r yBase Q cx2r yCy2 x9 yB2 L x9 yD2 Q hc yCb x2 yD2 L x2 yB2 ' +
				'Q cx2 yCy2 l yBase L wd8 yN Z',
		),
		p('M x3 yS3 L x3 yC3 Q xfm yFcy x2 yC2 Z', { fill: 'darkenLess' }),
		p('M x5 yS3 L x5 yC3 Q xfmr yFcy x9 yC2 Z', { fill: 'darkenLess' }),
		p('M x2 yC2 L x2 yB2 M x9 yC2 L x9 yB2', { fill: 'none' }),
	]);
}

export const EXACT_ELLIPSE_RIBBON_PRESET_DEFINITIONS: Record<
	string,
	PresetShapeGeometryDefinition
> = {
	ellipseRibbon: build('ellipseRibbon', false),
	ellipseRibbon2: build('ellipseRibbon2', true),
};
