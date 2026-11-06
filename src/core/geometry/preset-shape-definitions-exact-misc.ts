/**
 * Spec-exact geometry for presets that were wrong or missing their stroked
 * detail paths in the hand-written table: `wave`, `doubleWave`, `dodecagon`,
 * `can`, `cube`, `bevel`, `smileyFace`, `flowChartDocument`,
 * `flowChartPunchedTape` and `flowChartOnlineStorage`.
 *
 * Compared with PowerPoint 16 at a non-square aspect (outline and fill).
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const WAVE_Y = [
	'y1 */ h a1 100000',
	'dy2 */ y1 10 3',
	'y2 +- y1 0 dy2',
	'y3 +- y1 dy2 0',
	'y4 +- b 0 y1',
	'y5 +- y4 0 dy2',
	'y6 +- y4 dy2 0',
	'dx1 */ w a2 100000',
	'of2 */ w a2 50000',
];

const wave = exactDef(
	'wave',
	{ adj1: 12500, adj2: 0 },
	[
		'a1 pin 0 adj1 20000',
		'a2 pin -10000 adj2 10000',
		...WAVE_Y,
		'dx2 ?: of2 0 of2',
		'x2 +- l 0 dx2',
		'dx5 ?: of2 of2 0',
		'x5 +- r 0 dx5',
		'dx3 +/ dx2 x5 3',
		'x3 +- x2 dx3 0',
		'x4 +/ x3 x5 2',
		'x6 +- l dx5 0',
		'x10 +- r dx2 0',
		'x7 +- x6 dx3 0',
		'x8 +/ x7 x10 2',
		'it */ h a1 50000',
		'il max x2 x6',
		'ir min x5 x10',
		'ib +- b 0 it',
	],
	'il it ir ib',
	[p('M x2 y1 C x3 y2 x4 y3 x5 y1 L x10 y4 C x8 y6 x7 y5 x6 y4 Z')],
);

const doubleWave = exactDef(
	'doubleWave',
	{ adj1: 6250, adj2: 0 },
	[
		'a1 pin 0 adj1 12500',
		'a2 pin -10000 adj2 10000',
		...WAVE_Y,
		'dx2 ?: of2 0 of2',
		'x2 +- l 0 dx2',
		'dx8 ?: of2 of2 0',
		'x8 +- r 0 dx8',
		'dx3 +/ dx2 x8 6',
		'x3 +- x2 dx3 0',
		'dx4 +/ dx2 x8 3',
		'x4 +- x2 dx4 0',
		'x5 +/ x2 x8 2',
		'x6 +- x5 dx3 0',
		'x7 +/ x6 x8 2',
		'x9 +- l dx8 0',
		'x15 +- r dx2 0',
		'x10 +- x9 dx3 0',
		'x11 +- x9 dx4 0',
		'x12 +/ x9 x15 2',
		'x13 +- x12 dx3 0',
		'x14 +/ x13 x15 2',
		'it */ h a1 50000',
		'il max x2 x9',
		'ir min x8 x15',
		'ib +- b 0 it',
	],
	'il it ir ib',
	[
		p(
			'M x2 y1 C x3 y2 x4 y3 x5 y1 C x6 y2 x7 y3 x8 y1 L x15 y4 C x14 y6 x13 y5 x12 y4 C x11 y6 x10 y5 x9 y4 Z',
		),
	],
);

const dodecagon = exactDef(
	'dodecagon',
	undefined,
	[
		'x1 */ w 2894 21600',
		'x2 */ w 7906 21600',
		'x3 */ w 13694 21600',
		'x4 */ w 18706 21600',
		'y1 */ h 2894 21600',
		'y2 */ h 7906 21600',
		'y3 */ h 13694 21600',
		'y4 */ h 18706 21600',
	],
	'x1 y1 x4 y4',
	[p('M l y2 L x1 y1 L x2 t L x3 t L x4 y1 L r y2 L r y3 L x4 y4 L x3 b L x2 b L x1 y4 L l y3 Z')],
);

const can = exactDef(
	'can',
	{ adj: 25000 },
	[
		'maxAdj */ 50000 h ss',
		'a pin 0 adj maxAdj',
		'y1 */ ss a 200000',
		'y2 +- y1 y1 0',
		'y3 +- b 0 y1',
	],
	'l y2 r y3',
	[
		p('M l y1 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd2 Z', { stroke: false }),
		p('M l y1 A wd2 y1 cd2 21600000 Z', { fill: 'lighten', stroke: false }),
		p('M r y1 A wd2 y1 0 cd2 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd2 L l y1', { fill: 'none' }),
	],
);

const cube = exactDef(
	'cube',
	{ adj: 25000 },
	[
		'a pin 0 adj 100000',
		'y1 */ ss a 100000',
		'y4 +- b 0 y1',
		'y2 */ y4 1 2',
		'y3 +/ y1 b 2',
		'x4 +- r 0 y1',
		'x2 */ x4 1 2',
		'x3 +/ y1 r 2',
	],
	'l y1 x4 b',
	[
		p('M l y1 L y1 t L r t L r y4 L x4 b L l b Z', { stroke: false }),
		p('M x4 y1 L r t L r y4 L x4 b Z', { fill: 'darkenLess', stroke: false }),
		p('M l y1 L y1 t L r t L x4 y1 Z', { fill: 'lightenLess', stroke: false }),
		p('M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 M x4 y1 L r t M x4 y1 L x4 b', {
			fill: 'none',
		}),
	],
);

const bevel = exactDef(
	'bevel',
	{ adj: 12500 },
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
		p('M l t L r t L r b L l b Z', { stroke: false }),
		p('M l t L x1 x1 L x1 y2 L l b Z', { fill: 'lighten', stroke: false }),
		p('M r t L r b L x2 y2 L x2 x1 Z', { fill: 'darken', stroke: false }),
		p('M l t L r t L x2 x1 L x1 x1 Z', { fill: 'lightenLess', stroke: false }),
		p('M l b L x1 y2 L x2 y2 L r b Z', { fill: 'darkenLess', stroke: false }),
		p(
			'M l t L r t L r b L l b Z M x1 x1 L x1 y2 L x2 y2 L x2 x1 Z M l t L x1 x1 M r t L x2 x1 M r b L x2 y2 M l b L x1 y2',
			{ fill: 'none' },
		),
	],
);

const smileyFace = exactDef(
	'smileyFace',
	{ adj: 4653 },
	[
		'a pin -4653 adj 4653',
		'x1 */ w 4969 21699',
		'x2 */ w 6215 21600',
		'x3 */ w 13135 21600',
		'x4 */ w 16640 21600',
		'y1 */ h 7570 21600',
		'y3 */ h 16515 21600',
		'dy2 */ h a 100000',
		'y2 +- y3 0 dy2',
		'y4 +- y3 dy2 0',
		'dy3 */ h a 50000',
		'y5 +- y4 dy3 0',
		'idx cos wd2 2700000',
		'idy sin hd2 2700000',
		'il +- hc 0 idx',
		'ir +- hc idx 0',
		'it +- vc 0 idy',
		'ib +- vc idy 0',
		'wR */ w 1125 21600',
		'hR */ h 1125 21600',
		'ex1 +- x2 0 0',
		'ex2 +- x3 0 0',
	],
	'il it ir ib',
	[
		p('M l vc A wd2 hd2 cd2 21600000 Z', { stroke: false }),
		p('M ex1 y1 A wR hR cd2 21600000 Z M ex2 y1 A wR hR cd2 21600000 Z', {
			fill: 'darkenLess',
			stroke: false,
		}),
		p('M x1 y2 Q hc y5 x4 y2 Q hc y5 x1 y2 Z', { fill: 'none', stroke: false }),
		p(
			'M l vc A wd2 hd2 cd2 21600000 Z M ex1 y1 A wR hR cd2 21600000 Z M ex2 y1 A wR hR cd2 21600000 Z M x1 y2 Q hc y5 x4 y2',
			{ fill: 'none' },
		),
	],
);

const flowChartDocument = exactDef(
	'flowChartDocument',
	undefined,
	['ib */ h 17322 21600'],
	'l t r ib',
	[p('M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z', { w: 21600, h: 21600 })],
);

const flowChartPunchedTape = exactDef(
	'flowChartPunchedTape',
	undefined,
	['y1 */ h 1 10', 'y2 */ h 9 10'],
	'l y1 r y2',
	[
		p(
			'M l y1 A wd4 y1 cd2 -10800000 A wd4 y1 cd2 cd2 L r y2 A wd4 y1 0 -10800000 A wd4 y1 0 cd2 Z',
		),
	],
);

const flowChartOnlineStorage = exactDef(
	'flowChartOnlineStorage',
	undefined,
	['il */ w 1 6', 'ir */ w 5 6'],
	'il t ir b',
	[p('M 1 0 L 6 0 A 1 3 3cd4 -10800000 L 1 6 A 1 3 cd4 cd2 Z', { w: 6, h: 6 })],
);

export const EXACT_MISC_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	wave,
	doubleWave,
	dodecagon,
	can,
	cube,
	bevel,
	smileyFace,
	flowChartDocument,
	flowChartPunchedTape,
	flowChartOnlineStorage,
};
