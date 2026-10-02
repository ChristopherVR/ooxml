/** Shared helpers for the chart SVG painters. */

export interface Rect {
	x: number;
	y: number;
	w: number;
	h: number;
}

export const TEXT_COLOR = '#595959';
export const GRID_COLOR = '#D9D9D9';
export const AXIS_COLOR = '#BFBFBF';
export const FONT_SIZE = 10;

/** Escapes text for SVG content and attribute values. */
export function esc(text: string): string {
	return text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

/** A coordinate with at most two decimals. */
export const n = (value: number): string => String(Math.round(value * 100) / 100);

/** Rough text width for layout (no measurement available in a DOM-free painter). */
export const textWidth = (text: string, size = FONT_SIZE): number => text.length * size * 0.55;

export function text(
	x: number,
	y: number,
	content: string,
	attrs: {
		anchor?: 'start' | 'middle' | 'end';
		size?: number;
		weight?: string;
		baseline?: string;
		fill?: string;
	} = {},
): string {
	const parts = [`x="${n(x)}"`, `y="${n(y)}"`];
	parts.push(`font-size="${attrs.size ?? FONT_SIZE}"`);
	parts.push(`fill="${attrs.fill ?? TEXT_COLOR}"`);
	if (attrs.anchor && attrs.anchor !== 'start') parts.push(`text-anchor="${attrs.anchor}"`);
	if (attrs.weight) parts.push(`font-weight="${attrs.weight}"`);
	if (attrs.baseline) parts.push(`dominant-baseline="${attrs.baseline}"`);
	return `<text ${parts.join(' ')}>${esc(content)}</text>`;
}

export const line = (
	x1: number,
	y1: number,
	x2: number,
	y2: number,
	stroke: string,
	width = 1,
): string =>
	`<line x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" stroke="${stroke}" stroke-width="${width}"/>`;

export const rect = (x: number, y: number, w: number, h: number, fill: string): string =>
	`<rect x="${n(x)}" y="${n(y)}" width="${n(Math.max(0, w))}" height="${n(Math.max(0, h))}" fill="${fill}"/>`;

/** Shortens text to fit `maxWidth` (estimated) with an ellipsis. */
export function fit(content: string, maxWidth: number, size = FONT_SIZE): string {
	if (textWidth(content, size) <= maxWidth) return content;
	const chars = Math.max(1, Math.floor(maxWidth / (size * 0.55)) - 1);
	return `${content.slice(0, chars)}…`;
}
