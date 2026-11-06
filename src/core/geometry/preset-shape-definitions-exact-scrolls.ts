/**
 * Geometry of `verticalScroll` and `horizontalScroll`.
 *
 * Measured on PowerPoint 16. `ch = ss * adj / 100000` is the roll diameter.
 * `verticalScroll` has a roll along the top (a circle `ch` across at the left
 * end with a small spiral, a round cap at the right end) and a curl at the
 * bottom left; the paper hangs between them from `x = ch` to `x = r - ch`, with
 * a rounded bottom right corner. `horizontalScroll` is the same figure turned
 * about its diagonal.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

const SIZE = [
	'a pin 0 adj 25000',
	'ch */ ss a 100000',
	'ch2 */ ch 1 2',
	'ch4 */ ch 1 4',
	'x3 +- ch ch2 0',
	'x3b +- x3 ch2 0',
	'ch8 */ ch 1 8',
	'cy3 +- ch2 ch4 0',
];

const verticalScroll = exactDef(
	'verticalScroll',
	{ adj: 12500 },
	[...SIZE, 'x4 +- r 0 ch', 'x6 +- r 0 ch2', 'y3 +- b 0 ch', 'y4 +- b 0 ch2', 'xsw +- x3 ch4 ch8'],
	'ch ch x6 y4',
	[
		p(
			'M ch ch2 A ch2 ch2 10800000 5400000 L x6 t A ch2 ch2 16200000 10800000 L x4 ch ' +
				'L x4 y4 A ch2 ch2 0 5400000 L ch2 b A ch2 ch2 5400000 10800000 L ch y3 Z',
		),
		p('M l y4 A ch2 ch2 10800000 21600000 Z M xsw cy3 A ch4 ch4 0 21600000 Z', {
			fill: 'darkenLess',
			stroke: false,
		}),
		p(
			'M x3 t A ch2 ch2 16200000 10800000 L x4 ch M ch2 y4 L ch y4 ' +
				'M x3b ch2 L x3 ch2 A ch4 ch4 16200000 -10800000',
			{ fill: 'none' },
		),
	],
);

const horizontalScroll = exactDef(
	'horizontalScroll',
	{ adj: 12500 },
	[...SIZE, 'x4 +- b 0 ch', 'x6 +- b 0 ch2', 'y3 +- r 0 ch', 'y4 +- r 0 ch2', 'cyw +- x3 ch8 0'],
	'ch ch y4 x6',
	[
		p(
			'M ch2 ch A ch2 ch2 16200000 -5400000 L l x6 A ch2 ch2 10800000 -10800000 L ch x4 ' +
				'L y4 x4 A ch2 ch2 5400000 -5400000 L r ch2 A ch2 ch2 0 -10800000 L y3 ch Z',
		),
		p('M y4 t A ch2 ch2 16200000 21600000 Z M ch cyw A ch4 ch4 0 21600000 Z', {
			fill: 'darkenLess',
			stroke: false,
		}),
		p(
			'M l x3 A ch2 ch2 10800000 -10800000 L ch x4 M y4 ch2 L y4 ch ' +
				'M ch2 x3b L ch2 x3 A ch4 ch4 10800000 10800000',
			{ fill: 'none' },
		),
	],
);

export const EXACT_SCROLL_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> = {
	verticalScroll,
	horizontalScroll,
};
