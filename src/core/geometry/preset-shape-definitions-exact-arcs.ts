/**
 * Spec-exact geometry for the arc family (`arc`, `pie`, `chord`, `blockArc`),
 * `pieWedge`, `noSmoking` and `halfFrame`.
 *
 * ECMA-376 arc angles are VISUAL angles: the direction from the ellipse centre
 * to the point. On a circle that equals the parametric angle; on any
 * non-square preset it does not, so PowerPoint and the old hand-written
 * entries parted ways as soon as `adj` was not a
 * multiple of 90 degrees. The arc engine now honours visual angles, so these
 * definitions use the specification's own end point formulas (`cat2`, `sat2`).
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

/**
 * The visual-angle end points of an arc, as the specification writes them:
 * `cat2` / `sat2` give the ellipse point seen along the visual angle.
 */
function endPoint(name: string, ang: string, wR: string, hR: string): string[] {
	return [
		`${name}wt sin ${wR} ${ang}`,
		`${name}ht cos ${hR} ${ang}`,
		`${name}dx cat2 ${wR} ${name}ht ${name}wt`,
		`${name}dy sat2 ${hR} ${name}ht ${name}wt`,
	];
}

const ARC_ANGLES = [
	'stAng0 pin 0 adj1 21599999',
	'enAng0 pin 0 adj2 21599999',
	'sw11 +- enAng0 0 stAng0',
	'sw12 +- sw11 21600000 0',
	'swAng ?: sw11 sw11 sw12',
	...endPoint('s', 'stAng0', 'wd2', 'hd2'),
	'x1 +- hc sdx 0',
	'y1 +- vc sdy 0',
];

const arc = exactDef('arc', { adj1: 16200000, adj2: 0 }, ARC_ANGLES, 'l t r b', [
	p('M x1 y1 A wd2 hd2 stAng0 swAng L hc vc Z', { stroke: false }),
	p('M x1 y1 A wd2 hd2 stAng0 swAng', { fill: 'none' }),
]);

const pie = exactDef('pie', { adj1: 0, adj2: 16200000 }, ARC_ANGLES, 'l t r b', [
	p('M x1 y1 A wd2 hd2 stAng0 swAng L hc vc Z'),
]);

const chord = exactDef('chord', { adj1: 2700000, adj2: 16200000 }, ARC_ANGLES, 'l t r b', [
	p('M x1 y1 A wd2 hd2 stAng0 swAng Z'),
]);

const blockArc = exactDef(
	'blockArc',
	{ adj1: 10800000, adj2: 0, adj3: 25000 },
	[
		'stAng0 pin 0 adj1 21599999',
		'istAng0 pin 0 adj2 21599999',
		'a3 pin 0 adj3 50000',
		'dr */ ss a3 100000',
		'iwd2 +- wd2 0 dr',
		'ihd2 +- hd2 0 dr',
		'sw11 +- istAng0 0 stAng0',
		'sw12 +- sw11 21600000 0',
		'swAng ?: sw11 sw11 sw12',
		'iswAng +- 0 0 swAng',
		...endPoint('s', 'stAng0', 'wd2', 'hd2'),
		...endPoint('e', 'istAng0', 'iwd2', 'ihd2'),
		'x1 +- hc sdx 0',
		'y1 +- vc sdy 0',
		'x2 +- hc edx 0',
		'y2 +- vc edy 0',
		'it */ ihd2 29289 100000',
	],
	'l t r b',
	[p('M x1 y1 A wd2 hd2 stAng0 swAng L x2 y2 A iwd2 ihd2 istAng0 iswAng Z')],
);

const pieWedge = exactDef(
	'pieWedge',
	undefined,
	['g1 cos w 2700000', 'g2 sin h 2700000'],
	'l t r b',
	[p('M l b A w h cd2 cd4 L r b Z')],
);

// noSmoking: a ring crossed by a bar along the box diagonal. The two holes are
// the inner ellipse minus the bar; the band edges meet the inner ellipse at
// parametric angles 45 +/- d and 225 +/- d degrees.
const noSmoking = exactDef(
	'noSmoking',
	{ adj: 18750 },
	[
		'a pin 0 adj 50000',
		'dr */ ss a 100000',
		'iwd2 +- wd2 0 dr',
		'ihd2 +- hd2 0 dr',
		'len mod iwd2 ihd2 0',
		'kk */ iwd2 ihd2 len',
		'den */ kk 282843 100000',
		's1 */ dr 100000 den',
		's2 */ s1 s1 100000',
		'c2 +- 100000 0 s2',
		'cc */ c2 100000 1',
		'c1 sqrt cc',
		'dd at2 c1 s1',
		'ta1 +- 2700000 dd 0',
		'ta2 +- 13500000 0 dd',
		'tb1 +- 13500000 dd 0',
		'tb2 +- 24300000 0 dd',
		'xa1 cos iwd2 ta1',
		'ya1 sin ihd2 ta1',
		'xa2 cos iwd2 ta2',
		'ya2 sin ihd2 ta2',
		'xb1 cos iwd2 tb1',
		'yb1 sin ihd2 tb1',
		'xb2 cos iwd2 tb2',
		'yb2 sin ihd2 tb2',
		'xa2p +- hc xa2 0',
		'ya2p +- vc ya2 0',
		'xb2p +- hc xb2 0',
		'yb2p +- vc yb2 0',
		'va1 at2 xa1 ya1',
		'va2 at2 xa2 ya2',
		'vb1 at2 xb1 yb1',
		'vb2 at2 xb2 yb2',
		'swa0 +- va1 0 va2',
		'swa1 +- swa0 0 21600000',
		'swa ?: swa0 swa1 swa0',
		'swb0 +- vb1 0 vb2',
		'swb1 +- swb0 0 21600000',
		'swb ?: swb0 swb1 swb0',
	],
	'l t r b',
	[
		p(
			'M l vc A wd2 hd2 cd2 21600000 Z M xa2p ya2p A iwd2 ihd2 va2 swa Z M xb2p yb2p A iwd2 ihd2 vb2 swb Z',
		),
	],
);

const halfFrame = exactDef(
	'halfFrame',
	{ adj1: 33333, adj2: 33333 },
	[
		'maxAdj2 */ 100000 w ss',
		'a2 pin 0 adj2 maxAdj2',
		'x1 */ ss a2 100000',
		'g1 */ h x1 w',
		'g2 +- h 0 g1',
		'maxAdj1 */ 100000 h ss',
		'a1 pin 0 adj1 maxAdj1',
		'y1 */ ss a1 100000',
		'dx2 */ y1 w h',
		'x2 +- r 0 dx2',
		'dy2 */ x1 h w',
		'y2 +- b 0 dy2',
		'cx1 */ x1 1 2',
		'cy1 +- y2 y1 0',
		'cx2 +- x2 x1 0',
		'cy2 */ y1 1 2',
	],
	'l t cx2 cy1',
	[p('M l t L r t L x2 y1 L x1 y1 L x1 y2 L l b Z')],
);

export const EXACT_ARC_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	arc,
	pie,
	chord,
	blockArc,
	pieWedge,
	noSmoking,
	halfFrame,
};
