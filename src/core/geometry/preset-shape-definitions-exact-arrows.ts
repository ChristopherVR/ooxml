/**
 * Spec-exact geometry for the multi-headed block arrows whose hand-written
 * entries disagreed with PowerPoint 16: `leftRightArrow`, `upDownArrow`,
 * `quadArrow`, `leftUpArrow` and `leftRightUpArrow`.
 *
 * Authored in the ECMA-376 `presetShapeDefinitions` layout and compared with
 * PowerPoint's raster at a non-square aspect with default and adjusted values.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const leftRightArrow = exactDef(
	'leftRightArrow',
	{ adj1: 50000, adj2: 50000 },
	[
		'maxAdj2 */ 100000 w ss',
		'a1 pin 0 adj1 100000',
		'a2 pin 0 adj2 maxAdj2',
		'x2 */ ss a2 100000',
		'x3 +- r 0 x2',
		'dy */ h a1 200000',
		'y1 +- vc 0 dy',
		'y2 +- vc dy 0',
		'dx1 */ y1 x2 hd2',
		'x1 +- x2 0 dx1',
		'x4 +- x3 dx1 0',
	],
	'x1 y1 x4 y2',
	[p('M l vc L x2 t L x2 y1 L x3 y1 L x3 t L r vc L x3 b L x3 y2 L x2 y2 L x2 b Z')],
);

const upDownArrow = exactDef(
	'upDownArrow',
	{ adj1: 50000, adj2: 50000 },
	[
		'maxAdj2 */ 50000 h ss',
		'a1 pin 0 adj1 100000',
		'a2 pin 0 adj2 maxAdj2',
		'y2 */ ss a2 100000',
		'y3 +- b 0 y2',
		'dx1 */ w a1 200000',
		'x1 +- hc 0 dx1',
		'x2 +- hc dx1 0',
		'dy1 */ x1 y2 wd2',
		'y1 +- y2 0 dy1',
		'y4 +- y3 dy1 0',
	],
	'x1 y1 x2 y4',
	[p('M l y2 L hc t L r y2 L x2 y2 L x2 y3 L r y3 L hc b L l y3 L x1 y3 L x1 y2 Z')],
);

const quadArrow = exactDef(
	'quadArrow',
	{ adj1: 22500, adj2: 22500, adj3: 22500 },
	[
		'a2 pin 0 adj2 50000',
		'maxAdj1 */ a2 2 1',
		'a1 pin 0 adj1 maxAdj1',
		'q1 +- 100000 0 maxAdj1',
		'maxAdj3 */ q1 1 2',
		'a3 pin 0 adj3 maxAdj3',
		'x1 */ ss a3 100000',
		'dx2 */ ss a2 100000',
		'x2 +- hc 0 dx2',
		'x5 +- hc dx2 0',
		'dx3 */ ss a1 200000',
		'x3 +- hc 0 dx3',
		'x4 +- hc dx3 0',
		'x6 +- r 0 x1',
		'y2 +- vc 0 dx2',
		'y5 +- vc dx2 0',
		'y3 +- vc 0 dx3',
		'y4 +- vc dx3 0',
		'y6 +- b 0 x1',
		'il +- x3 0 0',
	],
	'x3 y3 x4 y4',
	[
		p(
			'M l vc L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r vc L x6 y5 L x6 y4 L x4 y4 L x4 y6 L x5 y6 L hc b L x2 y6 L x3 y6 L x3 y4 L x1 y4 L x1 y5 Z',
		),
	],
);

const leftUpArrow = exactDef(
	'leftUpArrow',
	{ adj1: 25000, adj2: 25000, adj3: 25000 },
	[
		'a2 pin 0 adj2 50000',
		'maxAdj1 */ a2 2 1',
		'a1 pin 0 adj1 maxAdj1',
		'maxAdj3 +- 100000 0 maxAdj1',
		'a3 pin 0 adj3 maxAdj3',
		'x1 */ ss a3 100000',
		'dx2 */ ss a2 50000',
		'x2 +- r 0 dx2',
		'y2 +- b 0 dx2',
		'dx4 */ ss a2 100000',
		'x4 +- r 0 dx4',
		'y4 +- b 0 dx4',
		'dx3 */ ss a1 200000',
		'x3 +- x4 0 dx3',
		'x5 +- x4 dx3 0',
		'y3 +- y4 0 dx3',
		'y5 +- y4 dx3 0',
	],
	'l t x5 y5',
	[
		p(
			'M l y4 L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L x4 t L r x1 L x5 x1 L x5 y5 L x1 y5 L x1 b Z',
		),
	],
);

const leftRightUpArrow = exactDef(
	'leftRightUpArrow',
	{ adj1: 25000, adj2: 25000, adj3: 25000 },
	[
		'a2 pin 0 adj2 50000',
		'maxAdj1 */ a2 2 1',
		'a1 pin 0 adj1 maxAdj1',
		'q1 +- 100000 0 maxAdj1',
		'maxAdj3 */ q1 1 2',
		'a3 pin 0 adj3 maxAdj3',
		'x1 */ ss a3 100000',
		'dx2 */ ss a2 100000',
		'x2 +- hc 0 dx2',
		'x5 +- hc dx2 0',
		'dx3 */ ss a1 200000',
		'x3 +- hc 0 dx3',
		'x4 +- hc dx3 0',
		'x6 +- r 0 x1',
		'dy2 */ ss a2 50000',
		'y2 +- b 0 dy2',
		'y4 +- b 0 dx2',
		'y3 +- y4 0 dx3',
		'y5 +- y4 dx3 0',
	],
	'l x1 r y5',
	[
		p(
			'M l y4 L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r y4 L x6 b L x6 y5 L x1 y5 L x1 b Z',
		),
	],
);

export const EXACT_ARROW_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	leftRightArrow,
	upDownArrow,
	quadArrow,
	leftUpArrow,
	leftRightUpArrow,
};
