// Page Layout > Themes > Colors: a read-only swatch of the workbook theme palette.
import { THEME_SLOTS } from '@christophervr/xlsx-core';
import type { EditorContext } from '../../context.js';

const hex = (value: string | undefined): string =>
	`#${(value ?? '000000').replace(/[^0-9A-Fa-f]/g, '').slice(-6)}`;

/** One gallery item: an SVG strip of the theme's dark, light and six accent colours. */
export function themePaletteItems(
	ctx: EditorContext,
): { id: string; label: string; preview?: string }[] {
	const theme = ctx.workbook()?.theme;
	if (!theme) return [];
	const order = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9];
	const rects = order
		.map(
			(slot, i) =>
				`<rect x="${i * 6}" y="0" width="6" height="16" fill="${hex(theme.colors[slot])}"><title>${THEME_SLOTS[slot] ?? ''}</title></rect>`,
		)
		.join('');
	return [
		{
			id: 'current',
			label: ctx.t('Theme Colors'),
			preview: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 60 16" width="60" height="16">${rects}</svg>`,
		},
	];
}
