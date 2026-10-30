/**
 * Themed bevels. A shape whose BevelTopType, BevelTopWidth, BevelTopHeight, BevelContourSize or
 * BevelContourColor cells are `Themed` takes them from the `a:sp3d` of the effect style its Quick
 * Style effects matrix selects: `a:bevelT` (the twelve preset names are the gallery's order, so
 * Visio's type is the index plus one), `contourW` and `contourClr`. The Visio 16 themes that
 * ship carry `extrusionClr` and `contourClr` only, so for them nothing is drawn; a theme that
 * adds `bevelT` or `contourW` is read here (not verified against a Visio-authored theme).
 */
import type { VisioBevel } from './bevel-3d';
import type { Cells } from './sheet';
import { colorChoice, drawingPaint, themeChild } from './theme-color';
import { themeFormat, type ThemeResources } from './theme-resolve';

const PRESETS = [
	'circle',
	'relaxedInset',
	'cross',
	'coolSlant',
	'angle',
	'softRound',
	'convex',
	'slope',
	'divot',
	'riblet',
	'hardEdge',
	'artDeco',
];
const EMU = 914_400;
const measure = (node: Element, name: string, fallback: number): number => {
	const raw = node.getAttribute(name);
	return raw !== null && /^\d{1,10}$/.test(raw) ? Number(raw) / EMU : fallback;
};

export interface ThemeSolid {
	bevel?: VisioBevel;
	/** Contour width in inches and its colour, when the theme gives one. */
	contour?: { size: number; color?: string };
}
export function themeSolid(
	cells: Cells,
	resources: ThemeResources,
	placeholder: string,
): ThemeSolid | undefined {
	const solid = themeChild(themeFormat(cells, 'Effects', resources), 'sp3d');
	if (!solid) return undefined;
	const result: ThemeSolid = {};
	const top = themeChild(solid, 'bevelT');
	if (top) {
		const type = PRESETS.indexOf(top.getAttribute('prst') ?? 'circle') + 1;
		if (type > 0)
			result.bevel = {
				type,
				width: measure(top, 'w', 76_200 / EMU),
				height: measure(top, 'h', 76_200 / EMU),
			};
	}
	const size = measure(solid, 'contourW', 0);
	if (size > 0) {
		const paint = drawingPaint(
			colorChoice(themeChild(solid, 'contourClr')),
			new Map(),
			placeholder,
		);
		result.contour = { size, ...(paint ? { color: paint.color } : {}) };
	}
	return result;
}
