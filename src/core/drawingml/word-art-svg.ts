import { buildXml } from '../xml/index';
import { appendSvg, createSvg, setSvgAttributes } from './svg-elements';
import { WARP_PATH_GENERATORS } from './text-warp-path';
import { WORD_ART_FONT, type WordArtStyle } from './word-art';
import { wordArtLayout, wordArtShadowGeometry } from './word-art-layout';

export function wordArtTransforms(): string[] {
	return ['textNoShape', ...Object.keys(WARP_PATH_GENERATORS)];
}

/** The same outer shadow in the style picker, worksheet drawing and document. */
export function wordArtCssShadow(style: WordArtStyle, scale = 1): string | undefined {
	const shadow = wordArtShadowGeometry(style, scale);
	if (!shadow) return undefined;
	const alpha = Math.round(shadow.opacity * 255)
		.toString(16)
		.padStart(2, '0');
	return `${shadow.dx}px ${shadow.dy}px ${shadow.blur}px #${shadow.color}${alpha}`;
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
	const { style, width, height, x = 0, y = 0, scale = 1 } = options;
	if (![x, y].every(Number.isFinite)) return '';
	const layout = wordArtLayout(options);
	if (!layout) return '';
	const id = options.id.replace(/[^A-Za-z0-9_-]/g, '_');
	const shadow = wordArtCssShadow(style, scale);
	const { gradient, fill, outline, size } = layout;
	const gradientId = `${id}-gradient`;
	const svg = createSvg();
	setSvgAttributes(svg, {
		x,
		y,
		width,
		height,
		viewBox: `0 0 ${width} ${height}`,
		'data-word-art-transform': layout.preset ?? 'textNoShape',
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
			stroke: `#${outline.color}`,
			'stroke-width': outline.width,
			'paint-order': 'stroke fill',
		});
	}
	if (shadow) group.setAttribute('style', `filter:drop-shadow(${shadow})`);
	layout.lines.forEach((line, index) => {
		if (line.kind === 'straight') {
			appendSvg(group, 'text', { x: line.x, y: line.y }).textContent = line.text;
			return;
		}
		const pathId = `${id}-${index}`;
		const warped = appendSvg(group, 'g', {
			transform: `translate(${line.offsetX} ${line.offsetY})`,
		});
		appendSvg(appendSvg(warped, 'defs'), 'path', { id: pathId, d: line.path });
		appendSvg(appendSvg(warped, 'text'), 'textPath', {
			href: `#${pathId}`,
			startOffset: '50%',
		}).textContent = line.text;
	});
	return buildXml(svg);
}
