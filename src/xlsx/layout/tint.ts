/**
 * Excel's colour tint (`<color theme="4" tint="0.39997"/>`). Excel converts the base colour to
 * HLS on the Windows integer scale (`HLSMAX` 240, as `ColorRGBToHLS` does), scales the luminance
 * (`L * (1 + tint)` to darken, `L * (1 - tint) + HLSMAX * tint` to lighten) and converts back
 * with the same integer arithmetic.
 *
 * Measured against 252 colours read back from Excel 16 through COM (see
 * `__fixtures__/excel-tints.txt`): 211 match exactly, the rest are within 3 per channel. The
 * standard palette tints of the theme colours (the "Lighter 40%" family) all match.
 */

const HLSMAX = 240;
const RGBMAX = 255;
const idiv = (a: number, b: number): number => Math.trunc(a / b);

export interface Hls {
	h: number;
	l: number;
	s: number;
}

/** RGB (0-255) to integer HLS (0-240), the classic `ColorRGBToHLS` algorithm. */
export function rgbToHls240(r: number, g: number, b: number): Hls {
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = idiv((max + min) * HLSMAX + RGBMAX, 2 * RGBMAX);
	if (max === min) return { h: idiv(HLSMAX * 2, 3), l, s: 0 };
	const sum = max + min;
	const s =
		l <= HLSMAX / 2
			? idiv((max - min) * HLSMAX + idiv(sum, 2), sum)
			: idiv((max - min) * HLSMAX + idiv(2 * RGBMAX - sum, 2), 2 * RGBMAX - sum);
	const delta = (c: number): number =>
		idiv((max - c) * idiv(HLSMAX, 6) + idiv(max - min, 2), max - min);
	const rd = delta(r);
	const gd = delta(g);
	const bd = delta(b);
	let h: number;
	if (r === max) h = bd - gd;
	else if (g === max) h = idiv(HLSMAX, 3) + rd - bd;
	else h = idiv(2 * HLSMAX, 3) + gd - rd;
	if (h < 0) h += HLSMAX;
	if (h > HLSMAX) h -= HLSMAX;
	return { h, l, s };
}

function hueToRgb(n1: number, n2: number, hue: number): number {
	let h = hue;
	if (h < 0) h += HLSMAX;
	if (h > HLSMAX) h -= HLSMAX;
	const sixth = idiv(HLSMAX, 6);
	if (h < sixth) return n1 + idiv((n2 - n1) * h + idiv(HLSMAX, 12), sixth);
	if (h < idiv(HLSMAX, 2)) return n2;
	if (h < idiv(HLSMAX * 2, 3))
		return n1 + idiv((n2 - n1) * (idiv(HLSMAX * 2, 3) - h) + idiv(HLSMAX, 12), sixth);
	return n1;
}

const clamp255 = (v: number): number => Math.max(0, Math.min(255, v));

/** Integer HLS (0-240) back to RGB (0-255), the classic `ColorHLSToRGB` algorithm. */
export function hls240ToRgb({ h, l, s }: Hls): [number, number, number] {
	if (s === 0) {
		const v = clamp255(Math.round((l * RGBMAX) / HLSMAX));
		return [v, v, v];
	}
	const m2 =
		l <= HLSMAX / 2
			? idiv(l * (HLSMAX + s) + idiv(HLSMAX, 2), HLSMAX)
			: l + s - idiv(l * s + idiv(HLSMAX, 2), HLSMAX);
	const m1 = 2 * l - m2;
	const channel = (hue: number): number =>
		clamp255(idiv(hueToRgb(m1, m2, hue) * RGBMAX + idiv(HLSMAX, 2), HLSMAX));
	return [channel(h + idiv(HLSMAX, 3)), channel(h), channel(h - idiv(HLSMAX, 3))];
}

/** Parses `RRGGBB` or `AARRGGBB` (alpha ignored), with or without `#`. */
export function parseHex(hex: string): [number, number, number] | undefined {
	const clean = hex.replace(/^#/, '');
	const rgb = clean.length === 8 ? clean.slice(2) : clean;
	if (!/^[0-9a-fA-F]{6}$/.test(rgb)) return undefined;
	return [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16)) as [number, number, number];
}

/** `#RRGGBB` (upper case). */
export const toHexColor = (rgb: readonly number[]): string =>
	`#${rgb
		.map((v) => clamp255(Math.round(v)).toString(16).padStart(2, '0'))
		.join('')
		.toUpperCase()}`;

/** Applies an Excel tint (-1..1) to an `RRGGBB` colour; returns `#RRGGBB`. */
export function applyTint(hex: string, tint: number | undefined): string {
	const rgb = parseHex(hex) ?? [0, 0, 0];
	if (!tint) return toHexColor(rgb);
	const t = Math.max(-1, Math.min(1, tint));
	const hls = rgbToHls240(...rgb);
	const l = t < 0 ? Math.trunc(hls.l * (1 + t)) : Math.round(hls.l * (1 - t) + HLSMAX * t);
	return toHexColor(hls240ToRgb({ ...hls, l: Math.max(0, Math.min(HLSMAX, l)) }));
}
