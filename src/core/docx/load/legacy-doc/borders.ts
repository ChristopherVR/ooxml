// Borders and shading: the .doc Brc and Shd values as WordprocessingML border and `w:shd` elements.
import type { DocBorder, DocShading } from '@christophervr/ole2/doc-brc';
import { el } from './xml';

const KNOWN_BORDER = new Set([
	'single',
	'thick',
	'double',
	'dotted',
	'dashed',
	'dotDash',
	'dotDotDash',
	'triple',
	'thinThickSmallGap',
	'thickThinSmallGap',
	'thinThickThinSmallGap',
	'thinThickMediumGap',
	'thickThinMediumGap',
	'thinThickThinMediumGap',
	'thinThickLargeGap',
	'thickThinLargeGap',
	'thinThickThinLargeGap',
	'wave',
	'doubleWave',
	'dashSmallGap',
	'dashDotStroked',
	'threeDEmboss',
	'threeDEngrave',
	'outset',
	'inset',
]);

/** One border edge as `<w:top .../>` and the like; `name` is the element name without prefix. */
export function borderXml(name: string, border: DocBorder | undefined): string {
	if (!border || border.style === 'none') return '';
	const val = KNOWN_BORDER.has(border.style) ? border.style : 'single';
	return el(`w:${name}`, {
		'w:val': val,
		'w:sz': Math.min(96, Math.max(2, Math.round(border.widthPt * 8))),
		'w:space': Math.round(border.spacePt),
		'w:color': border.color ?? 'auto',
		'w:shadow': border.shadow ? 'true' : undefined,
		'w:frame': border.frame ? 'true' : undefined,
	});
}

export const bordersEqual = (a: DocBorder | undefined, b: DocBorder | undefined): boolean =>
	JSON.stringify(a) === JSON.stringify(b);

// Word's ipat values 2 to 25 and their ST_Shd names.
const PATTERNS: Record<number, string> = {
	2: 'pct5',
	3: 'pct10',
	4: 'pct20',
	5: 'pct25',
	6: 'pct30',
	7: 'pct40',
	8: 'pct50',
	9: 'pct60',
	10: 'pct70',
	11: 'pct75',
	12: 'pct80',
	13: 'pct90',
	14: 'darkHorizontal',
	15: 'darkVertical',
	16: 'darkDiagonalDown',
	17: 'darkDiagonalUp',
	18: 'darkGrid',
	19: 'darkTrellis',
	20: 'lightHorizontal',
	21: 'lightVertical',
	22: 'lightDiagonalDown',
	23: 'lightDiagonalUp',
	24: 'lightGrid',
	25: 'lightTrellis',
};

/** `<w:shd>` for a shading, or an empty string for no shading. */
export function shadingXml(shading: DocShading | undefined): string {
	if (!shading) return '';
	if (shading.pattern === 1)
		return el('w:shd', {
			'w:val': 'clear',
			'w:color': 'auto',
			'w:fill': shading.foreground ?? shading.fill ?? 'auto',
		});
	const pattern = PATTERNS[shading.pattern];
	if (pattern)
		return el('w:shd', {
			'w:val': pattern,
			'w:color': shading.foreground ?? 'auto',
			'w:fill': shading.background ?? 'auto',
		});
	const fill = shading.fill ?? shading.background;
	return fill ? el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': fill }) : '';
}
