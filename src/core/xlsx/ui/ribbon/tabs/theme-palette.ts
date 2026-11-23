// Page Layout > Themes: gallery items for built-in Office themes, colour schemes and font schemes.
import {
	THEME_COLOR_PRESETS,
	THEME_FONT_PRESETS,
	THEME_PRESETS,
	THEME_SLOTS,
} from '../../../index';
import type { EditorContext } from '../../context';

const hex = (value: string | undefined): string =>
	`#${(value ?? '000000').replace(/[^0-9A-Fa-f]/g, '').slice(-6)}`;

const escapeXml = (value: string): string =>
	value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const SWATCH_ORDER = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9] as const;

function colorStripSvg(colors: readonly string[], title?: string): string {
	const rects = SWATCH_ORDER.map(
		(slot, i) =>
			`<rect x="${i * 6}" y="0" width="6" height="16" fill="${hex(colors[slot])}"><title>${escapeXml(THEME_SLOTS[slot] ?? '')}</title></rect>`,
	).join('');
	const label = title ? `<title>${escapeXml(title)}</title>` : '';
	return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 60 16" width="60" height="16">${label}${rects}</svg>`;
}

function themeTileSvg(colors: readonly string[], majorFont: string, title: string): string {
	const bg = hex(colors[0]);
	const fg = hex(colors[1]);
	const accents = [4, 5, 6, 7, 8, 9]
		.map(
			(slot, i) =>
				`<rect x="${4 + i * 9}" y="26" width="8" height="6" rx="1" fill="${hex(colors[slot])}"/>`,
		)
		.join('');
	return (
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 62 36" width="62" height="36">` +
		`<title>${escapeXml(title)}</title>` +
		`<rect x="0.5" y="0.5" width="61" height="35" rx="2" fill="${bg}" stroke="#CCCCCC"/>` +
		`<text x="6" y="20" font-size="14" font-family="${escapeXml(majorFont)}, sans-serif" fill="${fg}">Aa</text>` +
		`${accents}</svg>`
	);
}

/** Gallery items for Page Layout > Themes (complete built-in Office themes). */
export function themePresetItems(
	ctx: EditorContext,
): { id: string; label: string; preview?: string }[] {
	if (!ctx.workbook()) return [];
	return THEME_PRESETS.map((preset) => ({
		id: preset.id,
		label: preset.label,
		preview: themeTileSvg(preset.colors, preset.majorFont, ctx.t(preset.label)),
	}));
}

/** Gallery items for Page Layout > Themes > Colors (built-in Office colour schemes). */
export function themePaletteItems(
	ctx: EditorContext,
): { id: string; label: string; preview?: string }[] {
	if (!ctx.workbook()) return [];
	return THEME_COLOR_PRESETS.map((preset) => ({
		id: preset.id,
		label: preset.label,
		preview: colorStripSvg(preset.colors, ctx.t(preset.label)),
	}));
}

/** Gallery items for Page Layout > Themes > Fonts (built-in Office font schemes). */
export function themeFontItems(
	ctx: EditorContext,
): { id: string; label: string; preview?: string }[] {
	if (!ctx.workbook()) return [];
	return THEME_FONT_PRESETS.map((preset) => ({
		id: preset.id,
		label: preset.label,
		preview: `font-family: "${preset.majorFont}", sans-serif`,
	}));
}
