/**
 * Geometry of `ribbon` and `ribbon2` (the straight banners).
 *
 * `ribbon2` has its tails at the bottom: the centre banner spans the top down to
 * `y2`, each tail runs from `y2` to the bottom edge and folds up behind the
 * banner, the fold being two half ellipses `wd32` wide and `hR` tall. `ribbon` is
 * the same shape turned over. Each preset is a silhouette (`norm`) plus the two
 * fold faces (`darkenLess`).
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const SHARED = [
	'a1 pin 0 adj1 33333',
	'a2 pin 25000 adj2 75000',
	'x10 +- r 0 wd8',
	'dx2 */ w a2 200000',
	'x2 +- hc 0 dx2',
	'x9 +- hc dx2 0',
	'x3 +- x2 wd32 0',
	'x8 +- x9 0 wd32',
	'x5 +- x2 wd8 0',
	'x6 +- x9 0 wd8',
	'x4 +- x5 0 wd32',
	'x7 +- x6 wd32 0',
	'hR */ h a1 400000',
	'dy1 */ h a1 200000',
	'dy2 */ h a1 100000',
];

const AV = { adj1: 16667, adj2: 50000 };

const ribbon2 = exactDef(
	'ribbon2',
	AV,
	[
		...SHARED,
		'y1 +- b 0 dy1',
		'y2 +- b 0 dy2',
		'y4 +- t dy2 0',
		'y3 +/ y4 b 2',
		'y5 +- b 0 hR',
		'y6 +- y2 hR 0',
	],
	'x2 t x9 y2',
	[
		p(
			'M l b L wd8 y3 L l y4 L x2 y4 L x2 hR A wd32 hR 10800000 5400000 L x8 t ' +
				'A wd32 hR 16200000 5400000 L x9 y4 L r y4 L x10 y3 L r b L x7 b ' +
				'A wd32 hR 5400000 5400000 L x6 y2 L x5 y2 L x5 y5 A wd32 hR 0 5400000 Z',
		),
		p('M x2 y4 L x2 y6 M x9 y4 L x9 y6', { fill: 'none' }),
		p('M x3 y2 A wd32 hR 16200000 -10800000 L x4 y1 A wd32 hR 16200000 5400000 L x5 y2 Z', {
			fill: 'darkenLess',
		}),
		p('M x8 y2 A wd32 hR 16200000 10800000 L x7 y1 A wd32 hR 16200000 -5400000 L x6 y2 Z', {
			fill: 'darkenLess',
		}),
	],
);

const ribbon = exactDef(
	'ribbon',
	AV,
	[
		...SHARED,
		'y1 +- t dy1 0',
		'y2 +- t dy2 0',
		'y4 +- b 0 dy2',
		'y3 +/ t y4 2',
		'y5 +- t hR 0',
		'y6 +- y2 0 hR',
		'yb +- b 0 hR',
	],
	'x2 y2 x9 b',
	[
		p(
			'M l t L wd8 y3 L l y4 L x2 y4 L x2 yb A wd32 hR 10800000 -5400000 L x8 b ' +
				'A wd32 hR 5400000 -5400000 L x9 y4 L r y4 L x10 y3 L r t L x7 t ' +
				'A wd32 hR 16200000 -5400000 L x6 y2 L x5 y2 L x5 y5 A wd32 hR 0 -5400000 Z',
		),
		p('M x2 y4 L x2 y6 M x9 y4 L x9 y6', { fill: 'none' }),
		p('M x3 y2 A wd32 hR 5400000 10800000 L x4 y1 A wd32 hR 5400000 -5400000 L x5 y2 Z', {
			fill: 'darkenLess',
		}),
		p('M x8 y2 A wd32 hR 5400000 -10800000 L x7 y1 A wd32 hR 5400000 5400000 L x6 y2 Z', {
			fill: 'darkenLess',
		}),
	],
);

export const EXACT_RIBBON_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	ribbon,
	ribbon2,
};
