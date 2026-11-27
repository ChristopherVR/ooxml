/** Colour, transform and inheritance helpers for the SVG to PDF converter. */

/** Thrown when a page holds something the vector path cannot represent. */
export class PdfFallback extends Error {}

export interface Rgba {
	r: number;
	g: number;
	b: number;
	a: number;
}

const NAMED: Record<string, string> = {
	black: '000000',
	white: 'ffffff',
	red: 'ff0000',
	green: '008000',
	blue: '0000ff',
	yellow: 'ffff00',
	gray: '808080',
	grey: '808080',
	orange: 'ffa500',
	purple: '800080',
	silver: 'c0c0c0',
	navy: '000080',
	maroon: '800000',
	teal: '008080',
	lime: '00ff00',
	aqua: '00ffff',
	cyan: '00ffff',
	fuchsia: 'ff00ff',
	magenta: 'ff00ff',
	olive: '808000',
	brown: 'a52a2a',
	pink: 'ffc0cb',
};

const unit = (value: number) => Math.min(1, Math.max(0, value));

/** A CSS colour as 0..1 channels; `null` for none or transparent; throws on anything unknown. */
export function parseColor(value: string | null | undefined): Rgba | null {
	const text = (value ?? '').trim().toLowerCase();
	if (!text || text === 'none' || text === 'transparent') return null;
	let hex = /^#([\da-f]{3,8})$/.exec(text)?.[1] ?? NAMED[text];
	if (hex) {
		if (hex.length === 3 || hex.length === 4) hex = [...hex].map((c) => c + c).join('');
		if (hex.length === 6 || hex.length === 8)
			return {
				r: parseInt(hex.slice(0, 2), 16) / 255,
				g: parseInt(hex.slice(2, 4), 16) / 255,
				b: parseInt(hex.slice(4, 6), 16) / 255,
				a: hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1,
			};
	}
	const rgb = /^rgba?\(([^)]*)\)$/.exec(text)?.[1];
	if (rgb) {
		const parts = rgb.split(/[\s,/]+/).filter(Boolean);
		const channel = (part: string | undefined) =>
			part === undefined ? 0 : part.endsWith('%') ? parseFloat(part) / 100 : parseFloat(part) / 255;
		const alpha = parts[3];
		return {
			r: unit(channel(parts[0])),
			g: unit(channel(parts[1])),
			b: unit(channel(parts[2])),
			a:
				alpha === undefined
					? 1
					: unit(alpha.endsWith('%') ? parseFloat(alpha) / 100 : parseFloat(alpha)),
		};
	}
	throw new PdfFallback(`colour "${value}"`);
}

/** Transform list as 2x3 matrices applied left to right (SVG order). */
export function parseTransform(value: string | null): number[][] {
	if (!value) return [];
	const out: number[][] = [];
	const pattern = /(\w+)\s*\(([^)]*)\)/g;
	for (const match of value.matchAll(pattern)) {
		const v = match[2]!
			.split(/[\s,]+/)
			.filter(Boolean)
			.map(Number);
		if (v.some((item) => !Number.isFinite(item))) throw new PdfFallback('transform');
		switch (match[1]) {
			case 'matrix':
				if (v.length !== 6) throw new PdfFallback('transform');
				out.push(v);
				break;
			case 'translate':
				out.push([1, 0, 0, 1, v[0] ?? 0, v[1] ?? 0]);
				break;
			case 'scale':
				out.push([v[0] ?? 1, 0, 0, v[1] ?? v[0] ?? 1, 0, 0]);
				break;
			case 'rotate': {
				const a = ((v[0] ?? 0) * Math.PI) / 180,
					c = Math.cos(a),
					s = Math.sin(a);
				const [cx, cy] = [v[1] ?? 0, v[2] ?? 0];
				out.push([c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]);
				break;
			}
			case 'skewX':
				out.push([1, 0, Math.tan(((v[0] ?? 0) * Math.PI) / 180), 1, 0, 0]);
				break;
			case 'skewY':
				out.push([1, Math.tan(((v[0] ?? 0) * Math.PI) / 180), 0, 1, 0, 0]);
				break;
			default:
				throw new PdfFallback('transform');
		}
	}
	return out;
}

/** The attribute on the element or the nearest ancestor that sets it (inherited properties). */
export function inherited(element: Element, name: string): string | null {
	for (let node: Element | null = element, depth = 0; node && depth < 64; depth++) {
		const value = node.getAttribute(name);
		if (value !== null) return value;
		node = node.parentElement;
	}
	return null;
}

export const number = (value: string | null, fallback = 0): number => {
	const parsed = value === null ? NaN : Number.parseFloat(value);
	return Number.isFinite(parsed) ? parsed : fallback;
};

export const hrefOf = (element: Element): string | null =>
	element.getAttributeNS('http://www.w3.org/1999/xlink', 'href') ?? element.getAttribute('href');

/** The `id` inside `url(#id)`. */
export const urlId = (value: string | null): string | undefined =>
	value ? /^url\(\s*['"]?#([^'")\s]+)['"]?\s*\)$/.exec(value.trim())?.[1] : undefined;
