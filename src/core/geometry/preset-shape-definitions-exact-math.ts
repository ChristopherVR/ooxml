/**
 * Spec-exact geometry for the six math operator presets (`mathPlus`,
 * `mathMinus`, `mathMultiply`, `mathDivide`, `mathEqual`, `mathNotEqual`).
 *
 * The earlier hand-written entries drew the operators the full width of the
 * box; ECMA-376 insets every glyph to 73490/100000 of the box, and builds
 * `mathMultiply` from two diagonal bars and `mathNotEqual` from two bars and a
 * slanted slash. Every figure here was checked pixel for pixel against
 * PowerPoint 16 at a non-square aspect with default and non-default
 * adjustments.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const mathPlus = exactDef(
	'mathPlus',
	{ adj1: 23520 },
	[
		'a1 pin 0 adj1 73490',
		'dx1 */ w 73490 200000',
		'dy1 */ h 73490 200000',
		'dx2 */ ss a1 200000',
		'x1 +- hc 0 dx1',
		'x2 +- hc 0 dx2',
		'x3 +- hc dx2 0',
		'x4 +- hc dx1 0',
		'y1 +- vc 0 dy1',
		'y2 +- vc 0 dx2',
		'y3 +- vc dx2 0',
		'y4 +- vc dy1 0',
	],
	'x1 y2 x4 y3',
	[
		p(
			'M x1 y2 L x2 y2 L x2 y1 L x3 y1 L x3 y2 L x4 y2 L x4 y3 L x3 y3 L x3 y4 L x2 y4 L x2 y3 L x1 y3 Z',
		),
	],
);

const mathMinus = exactDef(
	'mathMinus',
	{ adj1: 23520 },
	[
		'a1 pin 0 adj1 100000',
		'dy1 */ h a1 200000',
		'dx1 */ w 73490 200000',
		'y1 +- vc 0 dy1',
		'y2 +- vc dy1 0',
		'x1 +- hc 0 dx1',
		'x2 +- hc dx1 0',
	],
	'x1 y1 x2 y2',
	[p('M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z')],
);

const mathEqual = exactDef(
	'mathEqual',
	{ adj1: 23520, adj2: 11760 },
	[
		'a1 pin 0 adj1 36745',
		'2a1 */ a1 2 1',
		'mAdj2 +- 100000 0 2a1',
		'a2 pin 0 adj2 mAdj2',
		'dy1 */ h a1 100000',
		'dy2 */ h a2 200000',
		'dx1 */ w 73490 200000',
		'y2 +- vc 0 dy2',
		'y3 +- vc dy2 0',
		'y1 +- y2 0 dy1',
		'y4 +- y3 dy1 0',
		'x1 +- hc 0 dx1',
		'x2 +- hc dx1 0',
	],
	'x1 y1 x2 y4',
	[p('M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z M x1 y3 L x2 y3 L x2 y4 L x1 y4 Z')],
);

const mathMultiply = exactDef(
	'mathMultiply',
	{ adj1: 23520 },
	[
		'a1 pin 0 adj1 51965',
		'th */ ss a1 100000',
		'a at2 w h',
		'sa sin 1 a',
		'ca cos 1 a',
		'ta tan 1 a',
		'dl mod w h 0',
		'rw */ dl 51965 100000',
		'lM +- dl 0 rw',
		'xM */ ca lM 2',
		'yM */ sa lM 2',
		'dxAM */ sa th 2',
		'dyAM */ ca th 2',
		'xA +- xM 0 dxAM',
		'yA +- yM dyAM 0',
		'xB +- xM dxAM 0',
		'yB +- yM 0 dyAM',
		'xBC +- hc 0 xB',
		'yBC */ xBC ta 1',
		'yC +- yBC yB 0',
		'xD +- r 0 xB',
		'xE +- r 0 xA',
		'yFE +- vc 0 yA',
		'xFE */ yFE 1 ta',
		'xF +- xE 0 xFE',
		'xL +- xA xFE 0',
		'yG +- b 0 yA',
		'yH +- b 0 yB',
		'yI +- b 0 yC',
	],
	'xA yB xE yH',
	[
		p(
			'M xA yA L xB yB L hc yC L xD yB L xE yA L xF vc L xE yG L xD yH L hc yI L xB yH L xA yG L xL vc Z',
		),
	],
);

const mathDivide = exactDef(
	'mathDivide',
	{ adj1: 23520, adj2: 5880, adj3: 11760 },
	[
		'a1 pin 1000 adj1 36745',
		'ma1 +- 0 0 a1',
		'ma3h +- 73490 ma1 0',
		'ma3w */ 36745 w h',
		'maxAdj3 min ma3h ma3w',
		'a3 pin 1000 adj3 maxAdj3',
		'm4a3 */ -4 a3 1',
		'maxAdj2 +- 73490 m4a3 a1',
		'a2 pin 0 adj2 maxAdj2',
		'dy1 */ h a1 200000',
		'yg */ h a2 100000',
		'rad */ h a3 100000',
		'dx1 */ w 73490 200000',
		'x1 +- hc 0 dx1',
		'x3 +- hc dx1 0',
		'y3 +- vc 0 dy1',
		'y4 +- vc dy1 0',
		'a +- yg rad 0',
		'y2 +- y3 0 a',
		'y1 +- y2 0 rad',
		'rad2 */ rad 2 1',
		'y5 +- b 0 y1',
		'y6 +- y5 0 rad2',
	],
	'x1 y3 x3 y4',
	[
		p(
			'M hc y1 A rad rad 16200000 21600000 Z M hc y6 A rad rad 16200000 21600000 Z M x1 y3 L x3 y3 L x3 y4 L x1 y4 Z',
		),
	],
);

const mathNotEqual = exactDef(
	'mathNotEqual',
	{ adj1: 23520, adj2: 6600000, adj3: 11760 },
	[
		'a1 pin 0 adj1 50000',
		'crAng pin 4200000 adj2 6600000',
		'2a1 */ a1 2 1',
		'maxAdj3 +- 100000 0 2a1',
		'a3 pin 0 adj3 maxAdj3',
		'dy1 */ h a1 100000',
		'dy2 */ h a3 200000',
		'dx1 */ w 73490 200000',
		'x1 +- hc 0 dx1',
		'x8 +- hc dx1 0',
		'y2 +- vc 0 dy2',
		'y3 +- vc dy2 0',
		'y1 +- y2 0 dy1',
		'y4 +- y3 dy1 0',
		'cadj2 +- crAng 0 5400000',
		'xadj2 tan hd2 cadj2',
		'len mod xadj2 hd2 0',
		'bhw */ len dy1 hd2',
		'xt0 +- hc xadj2 0',
		'bhw2 */ bhw 1 2',
		'xl0 +- xt0 0 bhw2',
		'xr0 +- xt0 bhw2 0',
		'd1 */ xadj2 y1 hd2',
		'd2 */ xadj2 y2 hd2',
		'd3 */ xadj2 y3 hd2',
		'd4 */ xadj2 y4 hd2',
		'dB */ xadj2 h hd2',
		'l1 +- xl0 0 d1',
		'l2 +- xl0 0 d2',
		'l3 +- xl0 0 d3',
		'l4 +- xl0 0 d4',
		'lb +- xl0 0 dB',
		'r1 +- xr0 0 d1',
		'r2 +- xr0 0 d2',
		'r3 +- xr0 0 d3',
		'r4 +- xr0 0 d4',
		'rb +- xr0 0 dB',
	],
	'x1 y1 x8 y4',
	[
		p(
			'M xl0 t L xr0 t L r1 y1 L x8 y1 L x8 y2 L r2 y2 L r3 y3 L x8 y3 L x8 y4 L r4 y4 L rb b L lb b L l4 y4 L x1 y4 L x1 y3 L l3 y3 L l2 y2 L x1 y2 L x1 y1 L l1 y1 Z',
		),
	],
);

export const EXACT_MATH_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	mathPlus,
	mathMinus,
	mathMultiply,
	mathDivide,
	mathEqual,
	mathNotEqual,
};
