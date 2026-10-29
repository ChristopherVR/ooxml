// Paints for worksheet shapes: DrawingML colours resolved against the workbook theme, the fill
// (solid, gradient, or the colour an `xdr:style` reference names) and the outline.
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import { drawingColorCss } from '../../drawingml/drawing-color-css';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import { drawingSchemeColors } from '../drawing-theme';
import type { ThemePalette } from '../model';
import type { ShapeNode } from '../model-shapes';

export const EMU_PER_PX = 9525;

const DASHES: Record<string, number[]> = {
	dash: [4, 3],
	sysDash: [3, 1],
	dot: [1, 2],
	sysDot: [1, 1],
	dashDot: [4, 3, 1, 3],
	lgDash: [8, 3],
	lgDashDot: [8, 3, 1, 3],
	lgDashDotDot: [8, 3, 1, 3, 1, 3],
	sysDashDot: [3, 1, 1, 1],
	sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** Width in EMU of the outline a style `lnRef` index names (the Office default format scheme). */
const STYLE_LINE_EMU = [0, 9525, 12700, 19050];

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
				`<stop offset="${at(stop.position / 1000)}%" stop-color="${shapeColor(stop.color, theme) ?? '#808080'}"/>`,
		)
		.join('');
	return `<linearGradient id="${id}" x1="${at(0.5 - dx)}" y1="${at(0.5 - dy)}" x2="${at(0.5 + dx)}" y2="${at(0.5 + dy)}">${stops}</linearGradient>`;
}

/**
 * The fill of a shape: its own `spPr` fill, else the colour of the style's `fillRef` (index 0
 * means no fill, as for a text box), else none.
 */
export function shapeFill(node: ShapeNode, id: string, theme?: ThemePalette): ShapePaint {
	const fill = node.fill;
	if (fill?.kind === 'none') return { fill: 'none', defs: '' };
	if (fill?.kind === 'solid') return { fill: shapeColor(fill.color, theme) ?? 'none', defs: '' };
	if (fill?.kind === 'gradient' && fill.stops.length >= 2)
		return { fill: `url(#${id})`, defs: gradientDef(id, fill, theme) };
	if (fill?.kind === 'gradient' && fill.stops[0])
		return { fill: shapeColor(fill.stops[0].color, theme) ?? 'none', defs: '' };
	if (fill?.kind === 'pattern')
		return { fill: shapeColor(fill.foreground, theme) ?? 'none', defs: '' };
	if (fill) return { fill: 'none', defs: '' };
	const ref = node.style?.fill;
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
export function shapeStroke(node: ShapeNode, scale: number, theme?: ThemePalette): ShapeStroke {
	const line = node.line;
	const ref = node.style?.line;
	let stroke = 'none';
	if (line?.fill?.kind === 'solid') stroke = shapeColor(line.fill.color, theme) ?? 'none';
	else if (line?.fill?.kind === 'gradient' && line.fill.stops[0])
		stroke = shapeColor(line.fill.stops[0].color, theme) ?? 'none';
	else if (!line?.fill && ref && ref.idx > 0) stroke = shapeColor(ref.color, theme) ?? 'none';
	const styleWidth = STYLE_LINE_EMU[ref?.idx ?? 0] ?? STYLE_LINE_EMU[1]!;
	const emu = line?.widthEmu ?? (styleWidth || STYLE_LINE_EMU[1]!);
	const width = Math.max(0.5, (emu / EMU_PER_PX) * scale);
	const dash = line?.dash ? DASHES[line.dash] : undefined;
	return {
		stroke,
		width,
		...(dash ? { dash: dash.map((part) => Math.round(part * width * 100) / 100).join(' ') } : {}),
	};
}
