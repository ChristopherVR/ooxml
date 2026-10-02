// Pure translation of the core's CellView into CSS values (font shorthand, fills, pattern tiles,
// borders, alignment). The DOM painter only copies these onto recycled nodes.
import type {
	BordersView,
	EdgeView,
	FillView,
	FontView,
	PatternType,
} from '@christophervr/xlsx-core';

const FALLBACKS: Record<string, string> = {
	calibri: '"Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
	'calibri light': '"Calibri Light", "Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
	cambria: '"Cambria", "Caladea", Georgia, serif',
	arial: 'Arial, "Liberation Sans", Helvetica, sans-serif',
	'times new roman': '"Times New Roman", "Liberation Serif", Times, serif',
	'courier new': '"Courier New", "Liberation Mono", monospace',
	aptos: '"Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif',
};

/** A CSS font-family list for a workbook font name. */
export function cssFontFamily(name: string): string {
	const known = FALLBACKS[name.toLowerCase()];
	if (known) return known;
	const quoted = `"${name.replace(/["\\]/g, '')}"`;
	return `${quoted}, "Calibri", "Carlito", "Segoe UI", sans-serif`;
}

/** CSS `font` shorthand for a font at a zoom (percent). */
export function cssFont(font: FontView, zoom = 100): string {
	const size = Math.max(1, Math.round(font.sizePx * (zoom / 100) * 100) / 100);
	return `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${size}px ${cssFontFamily(font.family)}`;
}

export function textDecoration(font: FontView): string {
	const parts: string[] = [];
	if (font.underline) parts.push('underline');
	if (font.strike) parts.push('line-through');
	return parts.join(' ') || 'none';
}

// 4x4 tiles ('#' = pattern colour) approximating Excel's 8x8 pattern bitmaps.
const PATTERN_TILES: Partial<Record<PatternType, string[]>> = {
	gray0625: ['#...', '....', '....', '....'],
	gray125: ['#...', '....', '..#.', '....'],
	lightGray: ['#...', '..#.', '#...', '..#.'],
	mediumGray: ['#.#.', '.#.#', '#.#.', '.#.#'],
	darkGray: ['####', '#.#.', '####', '.#.#'],
	darkHorizontal: ['####', '####', '....', '....'],
	lightHorizontal: ['####', '....', '....', '....'],
	darkVertical: ['##..', '##..', '##..', '##..'],
	lightVertical: ['#...', '#...', '#...', '#...'],
	darkDown: ['##..', '.##.', '..##', '#..#'],
	lightDown: ['#...', '.#..', '..#.', '...#'],
	darkUp: ['..##', '.##.', '##..', '#..#'],
	lightUp: ['...#', '..#.', '.#..', '#...'],
	darkGrid: ['####', '##..', '####', '##..'],
	lightGrid: ['####', '#...', '#...', '#...'],
	darkTrellis: ['##.#', '.###', '#.##', '###.'],
	lightTrellis: ['#.#.', '.#..', '#.#.', '...#'],
};

const svgUrl = (svg: string): string => `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;

/** A repeating background image for a pattern fill, or undefined for solid / none. */
export function patternImage(pattern: PatternType, fg: string): string | undefined {
	const tile = PATTERN_TILES[pattern];
	if (!tile) return undefined;
	let rects = '';
	tile.forEach((line, y) => {
		for (let x = 0; x < line.length; x++)
			if (line[x] === '#') rects += `<rect x="${x}" y="${y}" width="1" height="1"/>`;
	});
	return svgUrl(
		`<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4" shape-rendering="crispEdges" fill="${fg}">${rects}</svg>`,
	);
}

export interface FillPaint {
	background?: string;
	image?: string;
}

export function fillPaint(fill: FillView | undefined): FillPaint | undefined {
	if (!fill) return undefined;
	if ('background' in fill) return { background: fill.background };
	if ('gradient' in fill) return { image: fill.gradient };
	if (fill.pattern === 'none') return undefined;
	if (fill.pattern === 'solid') return { background: fill.fg };
	const image = patternImage(fill.pattern, fill.fg);
	return image ? { background: fill.bg, image } : { background: fill.fg };
}

export const edgeCss = (edge: EdgeView | undefined): string =>
	edge
		? `${edge.style === 'double' ? Math.max(3, edge.widthPx) : edge.widthPx}px ${edge.style} ${edge.color}`
		: '';

export const hasBorders = (b: BordersView): boolean =>
	Boolean(b.top || b.right || b.bottom || b.left || b.diagonalUp || b.diagonalDown);

/** SVG for diagonal borders across a w x h cell. */
export function diagonalSvg(b: BordersView, w: number, h: number): string | undefined {
	if (!b.diagonalUp && !b.diagonalDown) return undefined;
	const line = (edge: EdgeView, x1: number, y1: number, x2: number, y2: number): string => {
		const dash =
			edge.style === 'dashed'
				? ' stroke-dasharray="4 2"'
				: edge.style === 'dotted'
					? ' stroke-dasharray="1 2"'
					: '';
		return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${edge.color}" stroke-width="${edge.widthPx}"${dash}/>`;
	};
	let body = '';
	if (b.diagonalDown) body += line(b.diagonalDown, 0, 0, w, h);
	if (b.diagonalUp) body += line(b.diagonalUp, 0, h, w, 0);
	return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${body}</svg>`;
}

/** CSS flex alignment for the text box. */
export function justify(hAlign: string): string {
	if (hAlign === 'right') return 'flex-end';
	if (hAlign === 'center' || hAlign === 'centerContinuous' || hAlign === 'distributed')
		return 'center';
	return 'flex-start';
}

export function alignItems(vAlign: string): string {
	if (vAlign === 'top') return 'flex-start';
	if (vAlign === 'center' || vAlign === 'distributed' || vAlign === 'justify') return 'center';
	return 'flex-end';
}

export function textAlign(hAlign: string): string {
	if (hAlign === 'right') return 'right';
	if (hAlign === 'center' || hAlign === 'centerContinuous') return 'center';
	if (hAlign === 'justify' || hAlign === 'distributed') return 'justify';
	return 'left';
}

/** Excel's `fill` alignment repeats the text across the cell width. */
export function repeatToFill(text: string, textWidth: number, cellWidth: number): string {
	if (!text || textWidth <= 0) return text;
	const times = Math.max(1, Math.floor(cellWidth / textWidth));
	return text.repeat(Math.min(times, 500));
}

/** Excel shows `#` across a cell when a number does not fit. */
export function hashes(cellInnerWidth: number, hashWidth: number): string {
	const n = Math.max(1, Math.floor(cellInnerWidth / Math.max(1, hashWidth)));
	return '#'.repeat(Math.min(n, 200));
}
