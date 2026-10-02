import type { Border, BorderEdge, BorderStyle, Fill, Font, ThemePalette } from '../model.js';
import { resolveColor } from './colors.js';
import type { BordersView, EdgeView, FillView, FontView } from './types.js';

const BLACK = '#000000';
const WHITE = '#FFFFFF';

/** CSS pixels per point at 96 dpi. */
const PX_PER_PT = 96 / 72;

/** `font` over `base`: every property `font` sets wins. */
export function mergeFont(base: Font, font: Font | undefined): Font {
	if (!font) return base;
	const out: Font = { ...base };
	for (const [key, value] of Object.entries(font) as [keyof Font, Font[keyof Font]][])
		if (value !== undefined) (out as Record<string, unknown>)[key] = value;
	return out;
}

/** A font ready to paint. Theme fonts (`scheme`) resolve through the workbook theme. */
export function fontView(font: Font, theme: ThemePalette): FontView {
	const family =
		font.scheme === 'major'
			? theme.majorFont || font.name || 'Calibri Light'
			: font.scheme === 'minor'
				? theme.minorFont || font.name || 'Calibri'
				: font.name || theme.minorFont || 'Calibri';
	const view: FontView = {
		family,
		sizePx: (font.size && font.size > 0 ? font.size : 11) * PX_PER_PT,
		bold: font.bold === true,
		italic: font.italic === true,
		strike: font.strike === true,
		color: resolveColor(font.color, theme, BLACK) ?? BLACK,
	};
	if (font.underline)
		view.underline =
			font.underline === 'double' || font.underline === 'doubleAccounting' ? 'double' : 'single';
	if (font.vertAlign === 'superscript') view.vertAlign = 'super';
	else if (font.vertAlign === 'subscript') view.vertAlign = 'sub';
	return view;
}

/**
 * A fill ready to paint. `differential` follows the dxf convention, where a solid fill's colour
 * is its `bgColor` (Excel writes conditional-format fills that way).
 */
export function fillView(
	fill: Fill | undefined,
	theme: ThemePalette,
	differential = false,
): FillView | undefined {
	if (!fill) return undefined;
	if (fill.type === 'gradient') {
		const stops = [...fill.stops]
			.sort((a, b) => a.position - b.position)
			.map((stop) => `${resolveColor(stop.color, theme, WHITE)} ${round(stop.position * 100)}%`);
		if (stops.length === 0) return undefined;
		if (stops.length === 1) stops.push(stops[0] ?? WHITE);
		const css =
			fill.gradient === 'path'
				? `radial-gradient(circle, ${stops.join(', ')})`
				: `linear-gradient(${round(((fill.degree ?? 0) + 90) % 360)}deg, ${stops.join(', ')})`;
		return { gradient: css };
	}
	if (fill.pattern === 'none') {
		// A dxf may give a bare background colour without a pattern type.
		if (differential && fill.bgColor)
			return { background: resolveColor(fill.bgColor, theme, WHITE) ?? WHITE };
		return undefined;
	}
	if (fill.pattern === 'solid') {
		const color = differential ? (fill.bgColor ?? fill.fgColor) : (fill.fgColor ?? fill.bgColor);
		return { background: resolveColor(color, theme, BLACK) ?? BLACK };
	}
	return {
		pattern: fill.pattern,
		fg: resolveColor(fill.fgColor, theme, BLACK) ?? BLACK,
		bg: resolveColor(fill.bgColor, theme, WHITE) ?? WHITE,
	};
}

const round = (n: number): number => Math.round(n * 100) / 100;

/** Pixel width and CSS-like style of each Excel border style at 100% zoom. */
export const BORDER_STYLES: Readonly<Record<BorderStyle, Omit<EdgeView, 'color'>>> = {
	hair: { widthPx: 1, style: 'dotted' },
	thin: { widthPx: 1, style: 'solid' },
	dotted: { widthPx: 1, style: 'dotted' },
	dashed: { widthPx: 1, style: 'dashed' },
	dashDot: { widthPx: 1, style: 'dashed' },
	dashDotDot: { widthPx: 1, style: 'dashed' },
	medium: { widthPx: 2, style: 'solid' },
	mediumDashed: { widthPx: 2, style: 'dashed' },
	mediumDashDot: { widthPx: 2, style: 'dashed' },
	mediumDashDotDot: { widthPx: 2, style: 'dashed' },
	slantDashDot: { widthPx: 2, style: 'dashed' },
	thick: { widthPx: 3, style: 'solid' },
	double: { widthPx: 3, style: 'double' },
};

export function edgeView(edge: BorderEdge | undefined, theme: ThemePalette): EdgeView | undefined {
	if (!edge) return undefined;
	const style = BORDER_STYLES[edge.style];
	if (!style) return undefined;
	return { ...style, color: resolveColor(edge.color, theme, BLACK) ?? BLACK };
}

/** `border` over `base` per edge (used for conditional-format borders). */
export function mergeBorder(base: Border, border: Border | undefined): Border {
	if (!border) return base;
	const out: Border = { ...base };
	for (const [key, value] of Object.entries(border) as [keyof Border, Border[keyof Border]][])
		if (value !== undefined) (out as Record<string, unknown>)[key] = value;
	return out;
}

export function bordersView(border: Border, theme: ThemePalette): BordersView {
	const out: BordersView = {};
	const set = (key: keyof BordersView, edge: BorderEdge | undefined): void => {
		const view = edgeView(edge, theme);
		if (view) out[key] = view;
	};
	set('top', border.top);
	set('right', border.right);
	set('bottom', border.bottom);
	set('left', border.left);
	if (border.diagonalUp) set('diagonalUp', border.diagonal);
	if (border.diagonalDown) set('diagonalDown', border.diagonal);
	return out;
}
