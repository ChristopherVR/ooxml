// Paints for worksheet shapes: DrawingML colours resolved against the workbook theme, the fill
// (solid, gradient, or the colour an `xdr:style` reference names) and the outline.
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import { drawingColorCss } from '../../drawingml/drawing-color-css';
import { capCompensatedDashArray, dashPatternUnits } from '../../drawingml/dash-pattern';
import {
	resolveThemeEffectStyle,
	resolveThemeFillStyle,
	resolveThemeLineStyle,
} from '../../drawingml/format-scheme';
import { officeFormatScheme } from '../office-format-scheme';
import { escapeAttr } from '../write/xml-out';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import { drawingSchemeColors } from '../drawing-theme';
import type { ThemePalette } from '../model';
import type { ShapeNode } from '../model-shapes';

export const EMU_PER_PX = 9525;

/** The colour as CSS against the workbook theme, or undefined when it cannot be resolved. */
export function shapeColor(
	color: DrawingColor | undefined,
	theme?: ThemePalette,
): string | undefined {
	if (!color) return undefined;
	const scheme = drawingSchemeColors(theme);
	return drawingColorCss(resolveDrawingColor(color, { scheme: (name) => scheme[name] }));
}

export interface ShapePaint {
	/** `none`, a CSS colour or `url(#id)` into `defs`. */
	fill: string;
	defs: string;
}

function gradientDef(
	id: string,
	fill: Extract<DrawingFill, { kind: 'gradient' }>,
	theme?: ThemePalette,
): string {
	const radians = ((fill.angle ?? 90) * Math.PI) / 180;
	const dx = Math.cos(radians) / 2;
	const dy = Math.sin(radians) / 2;
	const at = (v: number) => Math.round(v * 1000) / 1000;
	const stops = fill.stops
		.map(
			(stop) =>
				`<stop offset="${at(stop.position / 1000)}%" stop-color="${escapeAttr(shapeColor(stop.color, theme) ?? '#808080')}"/>`,
		)
		.join('');
	return `<linearGradient id="${escapeAttr(id)}" x1="${at(0.5 - dx)}" y1="${at(0.5 - dy)}" x2="${at(0.5 + dx)}" y2="${at(0.5 + dy)}">${stops}</linearGradient>`;
}

/**
 * The fill of a shape: its own `spPr` fill, else the colour of the style's `fillRef` (index 0
 * means no fill, as for a text box), else none.
 */
export function shapeFill(node: ShapeNode, id: string, theme?: ThemePalette): ShapePaint {
	const ref = node.style?.fill;
	const fill =
		node.fill ??
		(ref && resolveThemeFillStyle(theme?.formatScheme ?? officeFormatScheme(), ref.idx, ref.color));
	if (fill?.kind === 'none') return { fill: 'none', defs: '' };
	if (fill?.kind === 'solid') return { fill: shapeColor(fill.color, theme) ?? 'none', defs: '' };
	if (fill?.kind === 'gradient' && fill.stops.length >= 2)
		return { fill: `url(#${id})`, defs: gradientDef(id, fill, theme) };
	if (fill?.kind === 'gradient' && fill.stops[0])
		return { fill: shapeColor(fill.stops[0].color, theme) ?? 'none', defs: '' };
	if (fill?.kind === 'pattern')
		return { fill: shapeColor(fill.foreground, theme) ?? 'none', defs: '' };
	if (fill) return { fill: 'none', defs: '' };
	return {
		fill: ref && ref.idx > 0 ? (shapeColor(ref.color, theme) ?? 'none') : 'none',
		defs: '',
	};
}

export interface ShapeStroke {
	stroke: string;
	/** Width in pixels at the drawing scale. */
	width: number;
	dash?: string;
}

/** The outline: its own `a:ln`, else the style's `lnRef` colour (index 0 means no line). */
export function shapeStroke(
	node: Pick<ShapeNode, 'line' | 'style'>,
	scale: number,
	theme?: ThemePalette,
): ShapeStroke {
	const ref = node.style?.line;
	const inherited =
		ref && resolveThemeLineStyle(theme?.formatScheme ?? officeFormatScheme(), ref.idx, ref.color);
	const line = { ...inherited, ...node.line };
	let stroke = 'none';
	if (line?.fill?.kind === 'solid') stroke = shapeColor(line.fill.color, theme) ?? 'none';
	else if (line?.fill?.kind === 'gradient' && line.fill.stops[0])
		stroke = shapeColor(line.fill.stops[0].color, theme) ?? 'none';
	else if (!line?.fill && ref && ref.idx > 0) stroke = shapeColor(ref.color, theme) ?? 'none';
	const emu = line.widthEmu ?? 12700;
	const width = Math.max(0.5, (emu / EMU_PER_PX) * scale);
	const dash = dashPatternUnits(line?.dash);
	return {
		stroke,
		width,
		...(dash ? { dash: capCompensatedDashArray(dash, width) } : {}),
	};
}

export function shapeShadow(
	node: Pick<ShapeNode, 'shadow' | 'style' | 'hasExplicitEffects'>,
	theme?: ThemePalette,
) {
	const ref = node.style?.effect;
	if (node.shadow || node.hasExplicitEffects || !ref) return node.shadow;
	return resolveThemeEffectStyle(
		theme?.formatScheme ?? officeFormatScheme(),
		ref.idx,
		ref.color,
	)?.effects.find((effect) => effect.kind === 'outerShadow');
}
