// Maps BIFF8 formatting (XF, FONT, palette and XFEXT colours from the ole2 reader) onto the
// workbook model's resolved CellStyle. Palette colours become explicit RGB; XFEXT theme colours
// stay theme references; automatic colours are left unset so the default applies.
import type {
	XlsBorderEdge,
	XlsColor,
	XlsFont,
	XlsWorkbook,
	XlsXf,
} from '@christophervr/ole2/legacy-excel-workbook';
import type {
	Alignment,
	Border,
	BorderEdge,
	CellStyle,
	Color,
	Fill,
	Font,
	Protection,
} from '../model.js';

/** A model colour, or undefined for automatic/system colours. */
export function mapColor(color: XlsColor | undefined): Color | undefined {
	if (!color) return undefined;
	const tint = color.tint ? { tint: color.tint } : {};
	if (color.theme !== undefined) return { theme: color.theme, ...tint };
	if (color.rgb) return { rgb: `FF${color.rgb}`, ...tint };
	return undefined;
}

function mapFont(font: XlsFont | undefined, override: XlsColor | undefined): Font {
	if (!font) return {};
	const out: Font = { name: font.name, size: font.size };
	if (font.bold) out.bold = true;
	if (font.italic) out.italic = true;
	if (font.underline !== 'none') out.underline = font.underline;
	if (font.strike) out.strike = true;
	if (font.script !== 'none') out.vertAlign = font.script;
	if (font.family) out.family = font.family;
	const color = mapColor(override ?? font.color);
	if (color) out.color = color;
	return out;
}

function mapEdge(edge: XlsBorderEdge | undefined): BorderEdge | undefined {
	if (!edge) return undefined;
	const color = mapColor(edge.color);
	return color ? { style: edge.style, color } : { style: edge.style };
}

function mapBorder(xf: XlsXf): Border {
	const border: Border = {};
	for (const key of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) {
		const edge = mapEdge(xf.border[key]);
		if (edge) border[key] = edge;
	}
	if (border.diagonal) {
		if (xf.border.diagonalUp) border.diagonalUp = true;
		if (xf.border.diagonalDown) border.diagonalDown = true;
	}
	return border;
}

function mapFill(xf: XlsXf): Fill {
	if (xf.fill.pattern === 'none') return { type: 'pattern', pattern: 'none' };
	const fill: Fill = { type: 'pattern', pattern: xf.fill.pattern };
	const fg = mapColor(xf.fill.fg);
	const bg = mapColor(xf.fill.bg);
	if (fg) fill.fgColor = fg;
	if (bg) fill.bgColor = bg;
	return fill;
}

function mapAlignment(xf: XlsXf): Alignment | undefined {
	const source = xf.alignment;
	const out: Alignment = {};
	if (source.horizontal !== 'general') out.horizontal = source.horizontal;
	if (source.vertical !== 'bottom') out.vertical = source.vertical;
	if (source.wrap) out.wrapText = true;
	if (source.shrinkToFit) out.shrinkToFit = true;
	if (source.indent) out.indent = source.indent;
	if (source.rotation) out.textRotation = source.rotation;
	if (source.readingOrder) out.readingOrder = source.readingOrder;
	return Object.keys(out).length ? out : undefined;
}

function mapProtection(xf: XlsXf): Protection | undefined {
	const out: Protection = {};
	if (!xf.locked) out.locked = false;
	if (xf.hidden) out.hidden = true;
	return Object.keys(out).length ? out : undefined;
}

/** The resolved cell format of one BIFF8 XF. */
export function xfToCellStyle(xls: XlsWorkbook, xf: XlsXf): CellStyle {
	const style: CellStyle = {
		font: mapFont(xls.fonts[xf.font], xf.fontColor),
		fill: mapFill(xf),
		border: mapBorder(xf),
		numFmt: xf.numFmt,
	};
	const alignment = mapAlignment(xf);
	if (alignment) style.alignment = alignment;
	const protection = mapProtection(xf);
	if (protection) style.protection = protection;
	return style;
}

/** The default cell XF: 15 in every Excel-written file, else the first cell (non-style) XF. */
export function defaultXfIndex(xls: XlsWorkbook): number {
	if (xls.xfs[15] && !xls.xfs[15].isStyle) return 15;
	const index = xls.xfs.findIndex((xf) => !xf.isStyle);
	return index >= 0 ? index : 0;
}
