/**
 * Where WordArt's lines go and how they are painted, as `renderWordArtSvg` draws them: one font
 * size fitted to the box, each line centred on its warp path (or on a straight baseline), the
 * fill or two-stop gradient, the outline and the outer shadow. Shared by the SVG preview and
 * the PDF export, so both place the text the same way. DOM-free.
 */
import { EMU_PER_PIXEL } from '../units/constants';
import { buildWarpPath, WARP_PATH_GENERATORS } from './text-warp-path';
import {
	WORD_ART_OUTLINE_EMU,
	WORD_ART_SHADOW,
	normalizeWordArtGradient,
	wordArtRgb,
	type WordArtGradient,
	type WordArtStyle,
} from './word-art';

/** WordArt's outer shadow in pixels: offset, CSS blur radius (twice the deviation) and alpha. */
export interface WordArtShadowGeometry {
	color: string;
	dx: number;
	dy: number;
	blur: number;
	opacity: number;
}

export function wordArtShadowGeometry(
	style: WordArtStyle,
	scale = 1,
): WordArtShadowGeometry | undefined {
	const color = wordArtRgb(style.shadow);
	if (!color) return undefined;
	const offset = Math.SQRT1_2 * (WORD_ART_SHADOW.dist / EMU_PER_PIXEL) * scale;
	return {
		color,
		dx: offset,
		dy: offset,
		blur: (WORD_ART_SHADOW.blurRad / EMU_PER_PIXEL) * scale,
		opacity: WORD_ART_SHADOW.opacity / 100000,
	};
}

/** One line: centred on `path` (offset by `offsetX`, `offsetY`), or at `x` on baseline `y`. */
export type WordArtLayoutLine =
	| { text: string; kind: 'straight'; x: number; y: number }
	| { text: string; kind: 'path'; path: string; offsetX: number; offsetY: number };

export interface WordArtLayout {
	/** The fitted font size in px. */
	size: number;
	/** The warp preset drawn, or `undefined` for straight lines. */
	preset: string | undefined;
	lines: WordArtLayoutLine[];
	/** `RRGGBB`. */
	fill: string;
	gradient: WordArtGradient | undefined;
	outline: { color: string; width: number } | undefined;
	shadow: WordArtShadowGeometry | undefined;
}

export interface WordArtLayoutOptions {
	style: WordArtStyle;
	lines: readonly string[];
	width: number;
	height: number;
	fontPx: number;
	scale?: number;
}

/** The layout of `lines` in a `width` x `height` box; `undefined` when nothing can be drawn. */
export function wordArtLayout(options: WordArtLayoutOptions): WordArtLayout | undefined {
	const { style, lines, width, height, fontPx, scale = 1 } = options;
	if (![width, height, fontPx, scale].every(Number.isFinite)) return undefined;
	if (width <= 0 || height <= 0 || !lines.length) return undefined;
	const preset = style.warp?.preset;
	const usePath = preset !== undefined && Object.hasOwn(WARP_PATH_GENERATORS, preset);
	const adjustment = (name: string) => {
		const formula = style.warp?.guides?.find((guide) => guide.name === name)?.formula;
		const parsed = /^val\s+(-?\d+)$/u.exec(formula ?? '')?.[1];
		const value = Number(parsed);
		return Number.isFinite(value) ? value : undefined;
	};
	const longest = Math.max(1, ...lines.map((text) => Array.from(text).length));
	const size = Math.min(fontPx, (width - 4) / (longest * 0.61), height / (lines.length + 0.5));
	if (size <= 0) return undefined;
	const gap = size * 1.15;
	const inset = Math.min(width / 5, size * 0.5);
	const outline = wordArtRgb(style.outline);
	return {
		size,
		preset: usePath ? preset : undefined,
		fill: wordArtRgb(style.fill) ?? '26374A',
		gradient: normalizeWordArtGradient(style.gradient),
		outline: outline
			? { color: outline, width: (WORD_ART_OUTLINE_EMU / EMU_PER_PIXEL) * scale }
			: undefined,
		shadow: wordArtShadowGeometry(style, scale),
		lines: lines.map((text, index): WordArtLayoutLine => {
			if (!usePath)
				return {
					text,
					kind: 'straight',
					x: width / 2,
					y: height / 2 + size * 0.34 + (index - (lines.length - 1) / 2) * gap,
				};
			const top = size * 0.8;
			const path = buildWarpPath(
				preset,
				width - 2 * inset,
				Math.max(1, height - top - size * 0.26),
				index,
				lines.length,
				adjustment('adj') ?? adjustment('adj1'),
				adjustment('adj2'),
				gap,
			);
			return { text, kind: 'path', path, offsetX: inset, offsetY: top };
		}),
	};
}
