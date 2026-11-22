import { EMU_PER_PIXEL } from '../units/constants';
import { buildXml } from '../xml/index';
import { appendSvg, createSvg, setSvgAttributes } from './svg-elements';
import { buildWarpPath, WARP_PATH_GENERATORS } from './text-warp-path';
import {
	WORD_ART_FONT,
	WORD_ART_OUTLINE_EMU,
	WORD_ART_SHADOW,
	normalizeWordArtGradient,
	wordArtRgb,
	type WordArtStyle,
} from './word-art';

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
	if (![width, height, fontPx, x, y, scale].every(Number.isFinite)) return '';
	if (width <= 0 || height <= 0 || !lines.length) return '';
	const id = options.id.replace(/[^A-Za-z0-9_-]/g, '_');
	const outline = wordArtRgb(style.outline);
	const fill = wordArtRgb(style.fill) ?? '26374A';
	const shadow = wordArtCssShadow(style, scale);
	const gradient = normalizeWordArtGradient(style.gradient);
	const gradientId = `${id}-gradient`;
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
	const svg = createSvg();
	setSvgAttributes(svg, {
		x,
		y,
		width,
		height,
		viewBox: `0 0 ${width} ${height}`,
		'data-word-art-transform': usePath ? preset : 'textNoShape',
	});
	if (gradient) {
		const angle = (gradient.angle / 10800000) * Math.PI;
		const dx = Math.cos(angle) / 2;
		const dy = Math.sin(angle) / 2;
		const paint = appendSvg(appendSvg(svg, 'defs'), 'linearGradient', {
			id: gradientId,
			x1: 0.5 - dx,
			y1: 0.5 - dy,
			x2: 0.5 + dx,
			y2: 0.5 + dy,
		});
		appendSvg(paint, 'stop', { 'stop-color': `#${fill}` });
		appendSvg(paint, 'stop', { offset: 1, 'stop-color': `#${gradient.end}` });
	}
	const group = appendSvg(svg, 'g', {
		'text-anchor': 'middle',
		'font-family': `${WORD_ART_FONT}, sans-serif`,
		'font-size': size,
		'font-weight': 'bold',
		fill: gradient ? `url(#${gradientId})` : `#${fill}`,
	});
	if (outline) {
		setSvgAttributes(group, {
			stroke: `#${outline}`,
			'stroke-width': (WORD_ART_OUTLINE_EMU / EMU_PER_PIXEL) * scale,
			'paint-order': 'stroke fill',
		});
	}
	if (shadow) group.setAttribute('style', `filter:drop-shadow(${shadow})`);
	const gap = size * 1.15;
	const inset = Math.min(width / 5, size * 0.5);
	lines.forEach((line, index) => {
		if (!usePath) {
			const baseline = height / 2 + size * 0.34 + (index - (lines.length - 1) / 2) * gap;
			appendSvg(group, 'text', { x: width / 2, y: baseline }).textContent = line;
			return;
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
		const warped = appendSvg(group, 'g', { transform: `translate(${inset} ${top})` });
		appendSvg(appendSvg(warped, 'defs'), 'path', { id: pathId, d: path });
		appendSvg(appendSvg(warped, 'text'), 'textPath', {
			href: `#${pathId}`,
			startOffset: '50%',
		}).textContent = line;
	});
	return buildXml(svg);
}
