/**
 * Spec-exact geometry for presets whose hand-written entries disagreed with
 * PowerPoint 16: the three tab corners, brackets and braces (single and
 * paired), `moon`, `diagStripe` and `nonIsoscelesTrapezoid`.
 *
 * Each definition mirrors the ECMA-376 `presetShapeDefinitions` layout (an
 * `avLst`, a `gdLst` of one-line formulas, a `pathLst`) and was compared with
 * PowerPoint's own raster at a non-square aspect, default and adjusted.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const CD2 = 'cd2';
const CD4 = 'cd4';

// Tabs: the shape is four small corner pieces, not a frame. Each piece is
// `ss / 10` wide.
const TAB_GUIDES = ['x1 */ ss 1 10', 'x2 +- r 0 x1', 'y2 +- b 0 x1'];

const cornerTabs = exactDef('cornerTabs', undefined, TAB_GUIDES, 'l t r b', [
	p('M l t L x1 t L l x1 Z M r t L r x1 L x2 t Z M l b L l y2 L x1 b Z M r b L x2 b L r y2 Z'),
]);

const squareTabs = exactDef('squareTabs', undefined, TAB_GUIDES, 'l t r b', [
	p(
		'M l t L x1 t L x1 x1 L l x1 Z M x2 t L r t L r x1 L x2 x1 Z M l y2 L x1 y2 L x1 b L l b Z M x2 y2 L r y2 L r b L x2 b Z',
	),
]);

const plaqueTabs = exactDef('plaqueTabs', undefined, TAB_GUIDES, 'l t r b', [
	p(
		`M l x1 A x1 x1 ${CD4} -5400000 L l t Z M x2 t A x1 x1 ${CD2} -5400000 L r t Z M r y2 A x1 x1 3cd4 -5400000 L r b Z M x1 b A x1 x1 0 -5400000 L l b Z`,
	),
]);

const bracketGuides = [
	'maxAdj */ 50000 h ss',
	'a pin 0 adj maxAdj',
	'y1 */ ss a 100000',
	'y2 +- b 0 y1',
	'dx1 cos w 2700000',
	'dy1 sin y1 2700000',
	'il +- r 0 dx1',
	'it +- y1 0 dy1',
	'ib +- b dy1 y1',
];

const leftBracket = exactDef('leftBracket', { adj: 8333 }, bracketGuides, 'il it r ib', [
	p(`M r b A w y1 ${CD4} ${CD4} L l y1 A w y1 ${CD2} ${CD4}`, { stroke: false }),
	p(`M r b A w y1 ${CD4} ${CD4} L l y1 A w y1 ${CD2} ${CD4}`, { fill: 'none' }),
]);

const rightBracket = exactDef(
	'rightBracket',
	{ adj: 8333 },
	[
		'maxAdj */ 50000 h ss',
		'a pin 0 adj maxAdj',
		'y1 */ ss a 100000',
		'y2 +- b 0 y1',
		'dx1 cos w 2700000',
		'dy1 sin y1 2700000',
		'ir +- l dx1 0',
		'it +- y1 0 dy1',
		'ib +- b dy1 y1',
	],
	'l it ir ib',
	[
		p(`M l t A w y1 3cd4 ${CD4} L r y2 A w y1 0 ${CD4}`, { stroke: false }),
		p(`M l t A w y1 3cd4 ${CD4} L r y2 A w y1 0 ${CD4}`, { fill: 'none' }),
	],
);

const braceGuides = [
	'a2 pin 0 adj2 100000',
	'q1 +- 100000 0 a2',
	'q2 min q1 a2',
	'q3 */ q2 1 2',
	'maxAdj1 */ q3 h ss',
	'a1 pin 0 adj1 maxAdj1',
	'y1 */ ss a1 100000',
	'y3 */ h a2 100000',
	'y4 +- y3 y1 0',
	'y2 +- y3 0 y1',
	'y5 +- b 0 y1',
	'dx1 cos wd2 2700000',
	'dy1 sin y1 2700000',
	'ir +- l dx1 0',
	'il +- r 0 dx1',
	'it +- y1 0 dy1',
	'ib +- b 0 it',
];

const LEFT_BRACE = `M r b A wd2 y1 ${CD4} ${CD4} L hc y4 A wd2 y1 0 -5400000 A wd2 y1 ${CD4} -5400000 L hc y1 A wd2 y1 ${CD2} ${CD4}`;
const RIGHT_BRACE = `M l t A wd2 y1 3cd4 ${CD4} L hc y2 A wd2 y1 ${CD2} -5400000 A wd2 y1 3cd4 -5400000 L hc y5 A wd2 y1 0 ${CD4}`;

const leftBrace = exactDef('leftBrace', { adj1: 8333, adj2: 50000 }, braceGuides, 'il it r ib', [
	p(LEFT_BRACE, { stroke: false }),
	p(LEFT_BRACE, { fill: 'none' }),
]);

const rightBrace = exactDef('rightBrace', { adj1: 8333, adj2: 50000 }, braceGuides, 'l it ir ib', [
	p(RIGHT_BRACE, { stroke: false }),
	p(RIGHT_BRACE, { fill: 'none' }),
]);

const bracketPair = exactDef(
	'bracketPair',
	{ adj: 16667 },
	[
		'a pin 0 adj 50000',
		'x1 */ ss a 100000',
		'x2 +- r 0 x1',
		'y2 +- b 0 x1',
		'il */ x1 29289 100000',
		'ir +- r 0 il',
		'ib +- b 0 il',
	],
	'il il ir ib',
	[
		p(
			`M l x1 A x1 x1 ${CD2} ${CD4} L x2 t A x1 x1 3cd4 ${CD4} L r y2 A x1 x1 0 ${CD4} L x1 b A x1 x1 ${CD4} ${CD4} Z`,
			{ stroke: false },
		),
		p(
			`M x1 b A x1 x1 ${CD4} ${CD4} L l x1 A x1 x1 ${CD2} ${CD4} M x2 t A x1 x1 3cd4 ${CD4} L r y2 A x1 x1 0 ${CD4}`,
			{ fill: 'none' },
		),
	],
);

const BRACE_PAIR_LEFT = `M x2 b A x1 x1 ${CD4} ${CD4} L x1 y3 A x1 x1 0 -5400000 A x1 x1 ${CD4} -5400000 L x1 x1 A x1 x1 ${CD2} ${CD4}`;
const BRACE_PAIR_RIGHT = `M x3 t A x1 x1 3cd4 ${CD4} L x4 y1 A x1 x1 ${CD2} -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 ${CD4}`;

const bracePair = exactDef(
	'bracePair',
	{ adj: 8333 },
	[
		'a pin 0 adj 25000',
		'x1 */ ss a 100000',
		'x2 */ ss a 50000',
		'x3 +- r 0 x2',
		'x4 +- r 0 x1',
		'y1 +- vc 0 x1',
		'y3 +- vc x1 0',
		'y4 +- b 0 x1',
		'il */ x1 29289 100000',
		'ir +- r 0 il',
		'ib +- b 0 il',
	],
	'il il ir ib',
	[
		p(`${BRACE_PAIR_LEFT} L x3 t ${BRACE_PAIR_RIGHT.replace(/^M x3 t /, '')} Z`, {
			stroke: false,
		}),
		p(`${BRACE_PAIR_LEFT} ${BRACE_PAIR_RIGHT}`, { fill: 'none' }),
	],
);

// moon: the outer edge is half of the box ellipse; the inner edge is an
// ellipse similar to it (a circle in box-normalised space) through the two
// tips and the point `adj` of the way across the box at mid height.
const moon = exactDef(
	'moon',
	{ adj: 50000 },
	[
		'a pin 0 adj 87500',
		'q +- 100000 0 a',
		'q2 */ q q 100000',
		'q3 +- q2 100000 0',
		'rho */ q3 50000 q',
		'cxn +- a rho 0',
		'cx */ w cxn 100000',
		'rx */ w rho 100000',
		'ry */ hd2 rho 100000',
		'dx0 +- w 0 cx',
		'nhd2 +- 0 0 hd2',
		'stRaw at2 dx0 nhd2',
		'stWrap +- stRaw 21600000 0',
		'stAng ?: stRaw stRaw stWrap',
		'm2st */ stAng -2 1',
		'swAng +- 21600000 m2st 0',
		// Text rectangle, as measured on PowerPoint 16: the left edge touches the
		// outer ellipse at 45 degrees, the right edge is the box centre line.
		'g0 */ ss a 100000',
		'g1 +- wd2 0 g0',
		'g2 */ g1 1 2',
		'idx cos wd2 2700000',
		'rl +- hc 0 idx',
		'rt +- vc 0 g2',
		'rr +- wd2 0 0',
		'rb +- vc g2 0',
	],
	'rl rt rr rb',
	[p(`M r b A w hd2 ${CD4} ${CD2} A rx ry stAng swAng Z`)],
);

const diagStripe = exactDef(
	'diagStripe',
	{ adj: 50000 },
	[
		'a pin 0 adj 100000',
		'x2 */ w a 100000',
		'x1 */ x2 1 2',
		'x3 +- x2 r 0',
		'x4 */ x3 1 2',
		'y2 */ h a 100000',
		'y1 */ y2 1 2',
		'y3 +- y2 b 0',
		'y4 */ y3 1 2',
	],
	'l t x4 y4',
	[p('M l y2 L x2 t L r t L l b Z')],
);

const nonIsoscelesTrapezoid = exactDef(
	'nonIsoscelesTrapezoid',
	{ adj1: 25000, adj2: 25000 },
	[
		'maxAdj */ 50000 w ss',
		'a1 pin 0 adj1 maxAdj',
		'a2 pin 0 adj2 maxAdj',
		'x1 */ ss a1 200000',
		'x2 */ ss a1 100000',
		'x3 */ ss a2 100000',
		'x4 +- r 0 x3',
		'il */ wd3 a1 maxAdj',
		'it */ hd3 a1 maxAdj',
		'ir +- r 0 il',
	],
	'il it ir b',
	[p('M l b L x2 t L x4 t L r b Z')],
);

// The pointer gap is centred on the parametric direction PowerPoint measured
// (`sdx`/`sdy` scale the pointer offset into the unit circle); the arc engine
// takes visual angles, so the two gap ends are turned back into the visual
// angles of their own points with `at2`.
const wedgeEllipseCallout = exactDef(
	'wedgeEllipseCallout',
	{ adj1: -20833, adj2: 62500 },
	[
		'dxPos */ w adj1 100000',
		'dyPos */ h adj2 100000',
		'xPos +- hc dxPos 0',
		'yPos +- vc dyPos 0',
		'sdx */ dxPos hd2 wd2',
		'sdy */ dyPos wd2 hd2',
		'pang at2 sdx sdy',
		'stAng0 +- pang 660000 0',
		'enAng0 +- pang 0 660000',
		'dx1 cos wd2 stAng0',
		'dy1 sin hd2 stAng0',
		'x1 +- hc dx1 0',
		'y1 +- vc dy1 0',
		'dx2 cos wd2 enAng0',
		'dy2 sin hd2 enAng0',
		'stAng at2 dx1 dy1',
		'enAng at2 dx2 dy2',
		'swAng0 +- enAng 0 stAng',
		'swAng1 +- swAng0 21600000 0',
		'swAng ?: swAng0 swAng0 swAng1',
		'il */ wd2 29289 100000',
		'it */ hd2 29289 100000',
		'ir +- r 0 il',
		'ib +- b 0 it',
	],
	'il it ir ib',
	[p('M x1 y1 A wd2 hd2 stAng swAng L xPos yPos Z')],
);

const heart = exactDef(
	'heart',
	undefined,
	[
		'dx1 */ w 49 48',
		'dx2 */ w 10 48',
		'x1 +- hc 0 dx1',
		'x2 +- hc 0 dx2',
		'x3 +- hc dx2 0',
		'x4 +- hc dx1 0',
		'y1 +- t 0 hd3',
		'rl */ w 1 6',
		'rt */ h 1 4',
		'rr +- r 0 rl',
		'rb +- b 0 hd3',
	],
	'rl rt rr rb',
	[p('M hc hd4 C x3 y1 x4 hd4 hc b C x1 hd4 x2 y1 hc hd4 Z')],
);

const lightningBolt = exactDef(
	'lightningBolt',
	undefined,
	[
		'x1 */ w 5022 21600',
		'x2 */ w 11050 21600',
		'x3 */ w 8472 21600',
		'x4 */ w 12860 21600',
		'x5 */ w 14767 21600',
		'y1 */ h 3890 21600',
		'y2 */ h 6080 21600',
		'y3 */ h 8382 21600',
		'y4 */ h 14915 21600',
		'il */ w 8472 21600',
		'it */ h 8382 21600',
		'ir */ w 12222 21600',
		'ib */ h 13987 21600',
	],
	'il it ir ib',
	[
		p(
			'M 8472 0 L 12860 6080 L 11050 6797 L 16577 12007 L 14767 12877 L 21600 21600 L 10012 14915 L 12222 13987 L 5022 9705 L 7602 8382 L 0 3890 Z',
			{ w: 21600, h: 21600 },
		),
	],
);

const cloud = exactDef(
	'cloud',
	undefined,
	['il */ w 2977 21600', 'it */ h 3262 21600', 'ir */ w 17087 21600', 'ib */ h 17337 21600'],
	'il it ir ib',
	[
		p(
			'M 3900 14370 A 6753 9190 -11429249 7426832 A 5333 7267 -8646143 5396714 A 4365 5945 -8748475 5983381 A 4857 6595 -7859164 7034504 A 5333 7273 -4722533 6541615 A 6775 9220 -2776035 7816140 A 5785 7867 37501 6842000 A 6752 9215 1347096 6910353 A 7720 10543 3974558 4542661 A 4360 5918 -16496525 8804134 A 4345 5945 -14809710 9151131 Z',
			{ w: 43200, h: 43200 },
		),
	],
);

export const EXACT_BASIC_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	cornerTabs,
	squareTabs,
	plaqueTabs,
	leftBracket,
	rightBracket,
	leftBrace,
	rightBrace,
	bracketPair,
	bracePair,
	moon,
	diagStripe,
	nonIsoscelesTrapezoid,
	wedgeEllipseCallout,
	heart,
	lightningBolt,
	cloud,
};
