// Gallery data for Format as Table (the 60 built-in table style names with an SVG swatch) and
// Cell Styles (the workbook's named styles plus the core's built-in cell styles).
import {
	BUILTIN_CELL_STYLES,
	DEFAULT_THEME,
	resolveColor,
	type Fill,
	type Font,
	type ThemePalette,
	type Workbook,
} from '@christophervr/xlsx-core';

export interface TableStyleInfo {
	name: string;
	family: 'Light' | 'Medium' | 'Dark';
	index: number;
}

export const TABLE_STYLES: readonly TableStyleInfo[] = [
	...Array.from({ length: 21 }, (_v, i) => ({
		name: `TableStyleLight${i + 1}`,
		family: 'Light' as const,
		index: i + 1,
	})),
	...Array.from({ length: 28 }, (_v, i) => ({
		name: `TableStyleMedium${i + 1}`,
		family: 'Medium' as const,
		index: i + 1,
	})),
	...Array.from({ length: 11 }, (_v, i) => ({
		name: `TableStyleDark${i + 1}`,
		family: 'Dark' as const,
		index: i + 1,
	})),
];

/** The accent a built-in table style uses (`000000`-ish for the first of each group of seven). */
export function tableStyleColor(info: TableStyleInfo, theme: ThemePalette): string {
	const slot = (info.index - 1) % 7;
	if (slot === 0) return info.family === 'Dark' ? '404040' : '808080';
	return theme.colors[3 + slot] ?? '4472C4';
}

const escapeAttr = (s: string) => s.replace(/[^0-9A-Fa-f]/g, '');

/** A 40x30 SVG swatch: header row, banded body. */
export function tableStylePreview(info: TableStyleInfo, theme: ThemePalette): string {
	const c = `#${escapeAttr(tableStyleColor(info, theme))}`;
	const dark = info.family === 'Dark';
	const header = info.family === 'Light' ? '#ffffff' : c;
	const band = dark ? c : info.family === 'Medium' ? `${c}55` : `${c}33`;
	const body = dark ? '#3a3a3a' : '#ffffff';
	const rows = [0, 1, 2, 3]
		.map(
			(r) =>
				`<rect x="1" y="${7 + r * 5.5}" width="38" height="5.5" fill="${r % 2 ? body : band}" opacity="${dark && r % 2 ? 1 : 0.9}"/>`,
		)
		.join('');
	const headerLine =
		info.family === 'Light' ? `<path d="M1 7h38" stroke="${c}" stroke-width="1"/>` : '';
	return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 30" width="40" height="30"><rect x="0.5" y="0.5" width="39" height="29" fill="${body}" stroke="${c}" stroke-width="${info.family === 'Light' ? 0.5 : 0}"/><rect x="1" y="1" width="38" height="6" fill="${header}"/>${rows}${headerLine}</svg>`;
}

/**
 * Gallery items: the workbook's own named styles first (they win by name), then the core's
 * built-in styles. The id is the style name `applyCellStyle` takes.
 */
export function cellStyleItems(
	workbook: Workbook | undefined,
): { id: string; label: string; preview?: string }[] {
	const theme = workbook?.theme ?? DEFAULT_THEME;
	const css = (fill: Fill | undefined, font: Partial<Font> | undefined): string => {
		const background = fill?.type === 'pattern' ? resolveColor(fill.fgColor, theme) : undefined;
		const parts = [
			`background:${background ?? '#ffffff'}`,
			`color:${resolveColor(font?.color, theme) ?? '#000000'}`,
		];
		if (font?.bold) parts.push('font-weight:bold');
		if (font?.italic) parts.push('font-style:italic');
		return parts.join(';');
	};
	const own = (workbook?.namedStyles ?? []).map((s) => ({
		id: s.name,
		label: s.name,
		preview: css(s.style.fill, s.style.font),
	}));
	const names = new Set(own.map((o) => o.label.toLowerCase()));
	const builtin = BUILTIN_CELL_STYLES.filter((s) => !names.has(s.name.toLowerCase())).map((s) => ({
		id: s.name,
		label: s.name,
		preview: css(s.fill, s.font),
	}));
	return [...own, ...builtin];
}
