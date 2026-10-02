import { elements, type XmlElement } from '../../xml/index.js';
import type {
	Alignment,
	Border,
	BorderEdge,
	BorderStyle,
	Color,
	Fill,
	Font,
	GradientStop,
	HorizontalAlignment,
	PatternType,
	Protection,
	UnderlineStyle,
	VerticalAlignment,
} from '../model.js';
import { att, boolAttr, childVal, numAttr, xChildren, xFirst } from './xml-util.js';

/** Reads a `CT_Color`; `palette` resolves overridden legacy indexed colours to RGB. */
export function parseColor(
	element: XmlElement | undefined,
	palette?: readonly string[],
): Color | undefined {
	if (!element) return undefined;
	const color: Color = {};
	const rgb = att(element, 'rgb');
	const theme = numAttr(element, 'theme');
	const indexed = numAttr(element, 'indexed');
	if (rgb && /^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/.test(rgb)) color.rgb = rgb.toUpperCase();
	else if (theme !== undefined) color.theme = theme;
	else if (indexed !== undefined) {
		const override = palette?.[indexed];
		if (override) color.rgb = override;
		else color.indexed = indexed;
	} else if (boolAttr(element, 'auto')) color.auto = true;
	const tint = numAttr(element, 'tint');
	if (tint !== undefined && tint !== 0) color.tint = tint;
	return Object.keys(color).length ? color : undefined;
}

const UNDERLINES = new Set<UnderlineStyle>([
	'single',
	'double',
	'singleAccounting',
	'doubleAccounting',
]);

/**
 * Reads a `CT_Font` (cell fonts, `rPr` run properties with `rFont`, or a `dxf` font). For
 * differential fonts an explicit `val="0"` is kept as `false` because it switches a flag off.
 */
export function parseFont(
	element: XmlElement | undefined,
	palette: readonly string[] | undefined,
	differential = false,
): Font {
	const font: Font = {};
	if (!element) return font;
	const flag = (local: string): boolean | undefined => {
		const node = xFirst(element, local);
		if (!node) return undefined;
		const on = boolAttr(node, 'val', true);
		return on || differential ? on : undefined;
	};
	const name = childVal(element, 'name') ?? childVal(element, 'rFont');
	if (name) font.name = name;
	const size = numAttr(xFirst(element, 'sz'), 'val');
	if (size !== undefined) font.size = size;
	const bold = flag('b');
	if (bold !== undefined) font.bold = bold;
	const italic = flag('i');
	if (italic !== undefined) font.italic = italic;
	const strike = flag('strike');
	if (strike !== undefined) font.strike = strike;
	const u = xFirst(element, 'u');
	if (u) {
		const value = (att(u, 'val') ?? 'single') as UnderlineStyle | 'none';
		if (UNDERLINES.has(value as UnderlineStyle)) font.underline = value as UnderlineStyle;
	}
	const color = parseColor(xFirst(element, 'color'), palette);
	if (color) font.color = color;
	const vert = childVal(element, 'vertAlign');
	if (vert === 'superscript' || vert === 'subscript') font.vertAlign = vert;
	const family = numAttr(xFirst(element, 'family'), 'val');
	if (family !== undefined) font.family = family;
	const scheme = childVal(element, 'scheme');
	if (scheme === 'major' || scheme === 'minor') font.scheme = scheme;
	return font;
}

/**
 * Reads a `CT_Fill`. In a differential format Excel paints a solid fill with `bgColor`; that is
 * normalised to `fgColor` so cell and conditional fills look alike in the model.
 */
export function parseFill(
	element: XmlElement | undefined,
	palette: readonly string[] | undefined,
	differential = false,
): Fill {
	const gradient = xFirst(element, 'gradientFill');
	if (gradient) {
		const stops: GradientStop[] = [];
		for (const stop of xChildren(gradient, 'stop')) {
			const color = parseColor(xFirst(stop, 'color'), palette);
			if (color) stops.push({ position: numAttr(stop, 'position') ?? 0, color });
		}
		const fill: Fill = {
			type: 'gradient',
			gradient: att(gradient, 'type') === 'path' ? 'path' : 'linear',
			stops,
		};
		const degree = numAttr(gradient, 'degree');
		if (degree !== undefined) fill.degree = degree;
		return fill;
	}
	const pattern = xFirst(element, 'patternFill');
	const declared = att(pattern, 'patternType') as PatternType | undefined;
	const fg = parseColor(xFirst(pattern, 'fgColor'), palette);
	const bg = parseColor(xFirst(pattern, 'bgColor'), palette);
	if (differential && pattern && (declared === undefined || declared === 'solid')) {
		if (!fg && !bg) return { type: 'pattern', pattern: 'none' };
		const color = bg ?? fg;
		return color
			? { type: 'pattern', pattern: 'solid', fgColor: color }
			: { type: 'pattern', pattern: 'solid' };
	}
	const type = declared ?? 'none';
	if (type === 'none') return { type: 'pattern', pattern: 'none' };
	const fill: Fill = { type: 'pattern', pattern: type };
	if (fg) fill.fgColor = fg;
	if (bg && !(type === 'solid' && bg.indexed === 64 && Object.keys(bg).length === 1))
		fill.bgColor = bg;
	return fill;
}

function parseEdge(
	element: XmlElement | undefined,
	palette: readonly string[] | undefined,
): BorderEdge | undefined {
	const style = att(element, 'style') as BorderStyle | 'none' | undefined;
	if (!element || !style || style === 'none') return undefined;
	const edge: BorderEdge = { style };
	const color = parseColor(xFirst(element, 'color'), palette);
	if (color) edge.color = color;
	return edge;
}

export function parseBorder(
	element: XmlElement | undefined,
	palette: readonly string[] | undefined,
): Border {
	const border: Border = {};
	if (!element) return border;
	for (const side of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) {
		const edge = parseEdge(
			xFirst(element, side) ??
				(side === 'left'
					? xFirst(element, 'start')
					: side === 'right'
						? xFirst(element, 'end')
						: undefined),
			palette,
		);
		if (edge) border[side] = edge;
	}
	if (boolAttr(element, 'diagonalUp')) border.diagonalUp = true;
	if (boolAttr(element, 'diagonalDown')) border.diagonalDown = true;
	return border;
}

export function parseAlignment(element: XmlElement | undefined): Alignment | undefined {
	if (!element) return undefined;
	const alignment: Alignment = {};
	const horizontal = att(element, 'horizontal') as HorizontalAlignment | undefined;
	if (horizontal && horizontal !== 'general') alignment.horizontal = horizontal;
	const vertical = att(element, 'vertical') as VerticalAlignment | undefined;
	if (vertical && vertical !== 'bottom') alignment.vertical = vertical;
	if (boolAttr(element, 'wrapText')) alignment.wrapText = true;
	if (boolAttr(element, 'shrinkToFit')) alignment.shrinkToFit = true;
	const indent = numAttr(element, 'indent');
	if (indent) alignment.indent = indent;
	const rotation = numAttr(element, 'textRotation');
	if (rotation) alignment.textRotation = rotation;
	const order = numAttr(element, 'readingOrder');
	if (order) alignment.readingOrder = order;
	return Object.keys(alignment).length ? alignment : undefined;
}

export function parseProtection(element: XmlElement | undefined): Protection | undefined {
	if (!element) return undefined;
	const protection: Protection = {};
	const locked = boolAttr(element, 'locked');
	if (locked === false) protection.locked = false;
	if (boolAttr(element, 'hidden')) protection.hidden = true;
	return Object.keys(protection).length ? protection : undefined;
}

/** `<colors><indexedColors>` overrides the legacy palette; returns `RRGGBB`/`AARRGGBB` entries. */
export function parseIndexedPalette(styleSheet: XmlElement): string[] | undefined {
	const indexed = xFirst(xFirst(styleSheet, 'colors'), 'indexedColors');
	if (!indexed) return undefined;
	const colors = elements(indexed)
		.map((node) => att(node, 'rgb')?.toUpperCase() ?? '')
		.filter((value) => /^[0-9A-F]{6}([0-9A-F]{2})?$/.test(value));
	return colors.length ? colors : undefined;
}
