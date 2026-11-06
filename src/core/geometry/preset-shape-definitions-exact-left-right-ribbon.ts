/**
 * Geometry of `leftRightRibbon`: a left-pointing arrow in front of a
 * right-pointing arrow, the left body rolled into the right one by a scroll
 * curl `wd32` wide and `h * adj3 / 100000` tall at the middle.
 *
 * Measured on PowerPoint 16: the left arrow body is `a1` of the height thick and
 * sits `adj3` above the right body; the curl is two half ellipses of radii
 * `wd32` x `hR` (`hR` a quarter of the offset) that join the left body's top edge
 * (at `hc`) to the right body's top edge, one bulging right, one left.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const leftRightRibbon = exactDef(
	'leftRightRibbon',
	{ adj1: 50000, adj2: 50000, adj3: 16667 },
	[
		'a3 pin 0 adj3 33333',
		'maxAdj1 +- 100000 0 a3',
		'a1 pin 0 adj1 maxAdj1',
		'w1 +- wd2 0 wd32',
		'maxAdj2 */ 100000 w1 ss',
		'a2 pin 0 adj2 maxAdj2',
		'x1 */ ss a2 100000',
		'x4 +- r 0 x1',
		'dy1 */ h a1 200000',
		'dy2 */ h a3 -200000',
		'ly1 +- vc dy2 dy1',
		'ly2 +- vc dy2 0',
		'ly2b +- ly2 dy1 0',
		'ry3 +- vc 0 dy2',
		'ry3t +- ry3 0 dy1',
		'ry4 +- vc dy1 dy2',
		'hR */ h a3 400000',
		'xl +- hc 0 wd32',
		'xr +- hc wd32 0',
		'yc2 +- ly1 hR hR',
		'yc3 +- yc2 hR 0',
		'tipL +- vc dy2 0',
		'tipR +- vc 0 dy2',
		'dy22 */ dy2 2 1',
		'lhb +- b dy22 0',
		'rht2 +- t 0 dy22',
	],
	'x1 ly1 x4 ry4',
	[
		p(
			'M l tipL L x1 t L x1 ly1 L hc ly1 A wd32 hR 16200000 10800000 ' +
				'A wd32 hR 16200000 -10800000 L x4 ry3t L x4 rht2 L r tipR L x4 b L x4 ry4 L hc ry4 ' +
				'A wd32 hR 5400000 5400000 L xl ly2b L x1 ly2b L x1 lhb Z',
		),
		p('M xl yc3 L xl ly2b M xr yc2 L xr ry3t', { fill: 'none' }),
	],
);

export const EXACT_LEFT_RIGHT_RIBBON_PRESET_DEFINITIONS: Record<
	string,
	PresetShapeGeometryDefinition
> = { leftRightRibbon };
