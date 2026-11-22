// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Draws a shape as SVG from its preset geometry (the shared DrawingML geometry tables), fill and
// outline, with its text centred inside. The editor, Print Layout and tests all use this one drawing.
import {
	drawingColorCss,
	resolveDrawingColor,
	parseCustomGeometry,
	drawingPathData,
} from '../drawingml/index';
import type { DrawingColor, DrawingFill } from '../drawingml/index';
import { evaluatePresetShape, lookupPresetShape } from '../geometry/preset-shape-evaluator';
import { normalizeHexColor } from '../color/color-primitives';
import { diagramColorTheme } from './diagram-theme';
import {
	DEFAULT_SHAPE_FILL,
	DEFAULT_SHAPE_LINE,
	DEFAULT_SHAPE_LINE_EMU,
	isLinePreset,
	type ShapeSpec,
} from './shape-model';
import type { ThemeCatalog } from './theme-model';
import type { InlineImage } from './inline-model';
import { capCompensatedDashArray, dashPatternUnits } from '../drawingml/dash-pattern';
import { appendSvg, createSvg, setSvgAttributes } from '../drawingml/svg-elements';
import { buildXml, parseXml } from '../xml/index';
import { EMU_PER_PIXEL } from './units';

const round = (value: number) => Math.round(value * 100) / 100;

/** A colour as CSS, resolved through the document theme. */
export function shapeColorCss(color: DrawingColor, theme?: ThemeCatalog): string | undefined {
	return drawingColorCss(resolveDrawingColor(color, diagramColorTheme(theme)));
}

/** The paint of an outline: its solid colour, none, or the default style's darker accent. */
function linePaint(fill: DrawingFill | undefined, theme?: ThemeCatalog): string {
	if (fill?.kind === 'none') return 'none';
	if (fill?.kind === 'solid') return shapeColorCss(fill.color, theme) ?? DEFAULT_SHAPE_LINE;
	return DEFAULT_SHAPE_LINE;
}

/** The paint of a fill: a CSS colour, or `none`. A gradient takes its first stop. */
function fillPaint(fill: DrawingFill | undefined, theme?: ThemeCatalog): string {
	if (!fill) return DEFAULT_SHAPE_FILL;
	if (fill.kind === 'none') return 'none';
	if (fill.kind === 'solid') return shapeColorCss(fill.color, theme) ?? DEFAULT_SHAPE_FILL;
	if (fill.kind === 'gradient' && fill.stops[0])
		return shapeColorCss(fill.stops[0].color, theme) ?? DEFAULT_SHAPE_FILL;
	return DEFAULT_SHAPE_FILL;
}

export interface ShapeAppearance {
	fill: string;
	stroke: string;
	strokeWidthPx: number;
	dash?: string;
}

/** The paint a shape ends up with, with Word's default style behind whatever the shape sets. */
export function shapeAppearance(spec: ShapeSpec, theme?: ThemeCatalog): ShapeAppearance {
	const line = spec.line;
	const strokeWidthPx = round((line?.widthEmu ?? DEFAULT_SHAPE_LINE_EMU) / EMU_PER_PIXEL);
	const dashUnits = dashPatternUnits(line?.dash);
	return {
		fill: isLinePreset(spec.preset) ? 'none' : fillPaint(spec.fill, theme),
		stroke: linePaint(line?.fill, theme),
		strokeWidthPx,
		...(dashUnits ? { dash: capCompensatedDashArray(dashUnits, strokeWidthPx) } : {}),
	};
}

const luminance = (css: string): number => {
	const hex = /^#([0-9a-f]{6})/i.exec(css)?.[1];
	if (!hex) return 0.5;
	const channel = (i: number) => Number.parseInt(hex.slice(i, i + 2), 16) / 255;
	return 0.2126 * channel(0) + 0.7152 * channel(2) + 0.0722 * channel(4);
};

/** Text colour inside a shape: white on the default (and any dark) fill, black on a light one. */
export const shapeTextColor = (fill: string): string =>
	fill === 'none' || luminance(fill) > 0.6 ? '#000000' : '#FFFFFF';

/**
 * The colour of a shape's text where its runs set none: the `wps:style` font reference colour as
 * Word draws it, or, without one, white on the default and dark fills and black on light ones.
 */
export function shapeDefaultTextColor(spec: ShapeSpec, theme?: ThemeCatalog): string {
	const font = spec.fontColor && shapeColorCss(spec.fontColor, theme);
	return font && /^#[0-9a-f]{6}$/i.test(font) ? font : shapeTextColor(fillPaint(spec.fill, theme));
}

const FONT_PX = 14.67;

/** Breaks `lines` to fit `widthPx` with an average-width estimate (no measuring in the core). */
export function wrapShapeText(lines: readonly string[], widthPx: number): string[] {
	const perLine = Math.max(1, Math.floor(widthPx / (FONT_PX * 0.52)));
	return lines.flatMap((line) => {
		if (line.length <= perLine) return [line];
		const out: string[] = [];
		let current = '';
		for (const word of line.split(/\s+/)) {
			if (current && current.length + 1 + word.length > perLine) {
				out.push(current);
				current = word;
			} else current = current ? `${current} ${word}` : word;
		}
		return [...out, current];
	});
}

/** Standalone `<svg>` for the shape at the given size, or undefined for an unknown preset. */
export function shapeSvg(
	spec: ShapeSpec,
	widthPx: number,
	heightPx: number,
	theme?: ThemeCatalog,
): string | undefined {
	const w = Math.max(1, widthPx);
	const h = Math.max(1, heightPx);
	const customPaths = spec.customGeometryXml
		? parseCustomGeometry(
				parseXml(spec.customGeometryXml).documentElement,
				w * EMU_PER_PIXEL,
				h * EMU_PER_PIXEL,
			)
		: undefined;
	const geometry = customPaths
		? {
				paths: customPaths.map((path) => ({
					d: drawingPathData(path, w, h),
					fill: path.fill,
					stroke: path.stroke,
				})),
				fillNone: false,
				textRect: undefined,
			}
		: lookupPresetShape(spec.preset)
			? evaluatePresetShape(spec.preset, w, h, spec.adjustments)
			: undefined;
	if (!geometry) return undefined;
	const svg = createSvg();
	setSvgAttributes(svg, {
		width: round(w),
		height: round(h),
		viewBox: `0 0 ${round(w)} ${round(h)}`,
		style: 'overflow:visible',
		role: 'img',
		'aria-label': lookupPresetShape(spec.preset) ? spec.preset : 'Shape',
	});
	const transform = [
		spec.rotation ? `rotate(${Number(spec.rotation)} ${round(w / 2)} ${round(h / 2)})` : '',
		spec.flipH || spec.flipV
			? `translate(${spec.flipH ? w : 0} ${spec.flipV ? h : 0}) scale(${spec.flipH ? -1 : 1} ${spec.flipV ? -1 : 1})`
			: '',
	]
		.filter(Boolean)
		.join(' ');
	const group = appendSvg(svg, 'g', transform ? { transform } : {});
	const look = shapeAppearance(spec, theme);
	for (const path of geometry.paths) {
		appendSvg(group, 'path', {
			d: path.d,
			fill: geometry.fillNone || path.fill === 'none' ? 'none' : look.fill,
			stroke: path.stroke === false ? 'none' : look.stroke,
			'stroke-width': look.strokeWidthPx,
			'stroke-linejoin': 'miter',
			...(look.dash ? { 'stroke-dasharray': look.dash } : {}),
		});
	}
	const textRect = geometry.textRect ?? { l: 0, t: 0, r: w, b: h };
	const lines = spec.text?.some(Boolean)
		? wrapShapeText(spec.text, textRect.r - textRect.l - 14)
		: [];
	if (lines.length) {
		const text = appendSvg(group, 'text', {
			'font-family': 'Calibri,Arial,sans-serif',
			'font-size': FONT_PX,
			fill: normalizeHexColor(spec.textColor, shapeDefaultTextColor(spec, theme)),
			'text-anchor': 'middle',
			'dominant-baseline': 'central',
		});
		const lineHeight = FONT_PX * 1.2;
		const first = (textRect.t + textRect.b) / 2 - ((lines.length - 1) * lineHeight) / 2;
		lines.forEach((line, index) => {
			appendSvg(text, 'tspan', {
				x: round((textRect.l + textRect.r) / 2),
				y: round(first + index * lineHeight),
			}).textContent = line;
		});
	}
	return buildXml(svg);
}

/**
 * A drawing's shape as SVG at the drawing's size. When the drawing carries its text box content as
 * paragraphs, the shape is drawn without its own centred text: the paragraphs are laid over it.
 */
export function drawingShapeSvg(
	image: Pick<InlineImage, 'shape' | 'textBoxContent' | 'widthPx' | 'heightPx'>,
	theme?: ThemeCatalog,
): string | undefined {
	if (!image.shape) return undefined;
	const { text: _text, ...geometry } = image.shape;
	return shapeSvg(
		image.textBoxContent ? geometry : image.shape,
		image.widthPx,
		image.heightPx,
		theme,
	);
}
