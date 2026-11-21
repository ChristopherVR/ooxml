import { EMU_PER_PIXEL } from '../units/constants';
import { buildWarpPath, WARP_PATH_GENERATORS } from './text-warp-path';
import {
	WORD_ART_FONT,
	WORD_ART_OUTLINE_EMU,
	WORD_ART_SHADOW,
	wordArtRgb,
	type WordArtStyle,
} from './word-art';

const escaped = (text: string) =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export function wordArtTransforms(): string[] {
	return ['textNoShape', ...Object.keys(WARP_PATH_GENERATORS)];
}

/** The same outer shadow in the style picker, worksheet drawing and document. */
export function wordArtCssShadow(style: WordArtStyle, scale = 1): string | undefined {
	const color = wordArtRgb(style.shadow);
	if (!color) return undefined;
	const offset = Math.SQRT1_2 * (WORD_ART_SHADOW.dist / EMU_PER_PIXEL) * scale;
	const blur = (WORD_ART_SHADOW.blurRad / EMU_PER_PIXEL) * scale;
	const alpha = Math.round((WORD_ART_SHADOW.opacity / 100000) * 255)
		.toString(16)
		.padStart(2, '0');
	return `${offset}px ${offset}px ${blur}px #${color}${alpha}`;
}

export interface WordArtSvgOptions {
	id: string;
	style: WordArtStyle;
	lines: readonly string[];
	width: number;
	height: number;
	fontPx: number;
	x?: number;
	y?: number;
	scale?: number;
}

/** Renders a shared SVG preview. Native saves carry the preset and adjustment guides. */
export function renderWordArtSvg(options: WordArtSvgOptions): string {
	const { style, lines, width, height, fontPx, x = 0, y = 0, scale = 1 } = options;
	if (width <= 0 || height <= 0 || !lines.length) return '';
	const id = options.id.replace(/[^A-Za-z0-9_-]/g, '_');
	const outline = wordArtRgb(style.outline);
	const fill = wordArtRgb(style.fill) ?? '26374A';
	const shadow = wordArtCssShadow(style, scale);
	const stroke = outline
		? ` stroke="#${outline}" stroke-width="${(WORD_ART_OUTLINE_EMU / EMU_PER_PIXEL) * scale}" paint-order="stroke fill"`
		: '';
	const preset = style.warp?.preset;
	const usePath = preset && Object.hasOwn(WARP_PATH_GENERATORS, preset);
	const adjustment = (name: string) => {
		const formula = style.warp?.guides?.find((guide) => guide.name === name)?.formula;
		const parsed = /^val\s+(-?\d+)$/u.exec(formula ?? '')?.[1];
		const value = Number(parsed);
		return Number.isFinite(value) ? value : undefined;
	};
	const longest = Math.max(1, ...lines.map((text) => Array.from(text).length));
	const size = Math.min(fontPx, (width - 4) / (longest * 0.61), height / (lines.length + 0.5));
	if (size <= 0) return '';
	const gap = size * 1.15;
	const inset = Math.min(width / 5, size * 0.5);
	const body = lines.map((line, index) => {
		if (!usePath) {
			const baseline = height / 2 + size * 0.34 + (index - (lines.length - 1) / 2) * gap;
			return `<text x="${width / 2}" y="${baseline}">${escaped(line)}</text>`;
		}
		const pathId = `${id}-${index}`;
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
		return `<g transform="translate(${inset} ${top})"><defs><path id="${pathId}" d="${path}"/></defs><text><textPath href="#${pathId}" startOffset="50%">${escaped(line)}</textPath></text></g>`;
	});
	return `<svg xmlns="http://www.w3.org/2000/svg" x="${x}" y="${y}" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" data-word-art-transform="${usePath ? escaped(preset) : 'textNoShape'}"><g text-anchor="middle" font-family="${WORD_ART_FONT}, sans-serif" font-size="${size}" font-weight="bold" fill="#${fill}"${stroke}${shadow ? ` style="filter:drop-shadow(${shadow})"` : ''}>${body.join('')}</g></svg>`;
}
