// `RRGGBB` / `AARRGGBB` parsing and upper-case `#RRGGBB` formatting on 0-255 channels.

/** Parses `RRGGBB` or `AARRGGBB` (alpha ignored), with or without `#`. */
export function parseArgbHex(hex: string): [number, number, number] | undefined {
	const clean = hex.replace(/^#/, '');
	const rgb = clean.length === 8 ? clean.slice(2) : clean;
	if (!/^[0-9a-fA-F]{6}$/.test(rgb)) return undefined;
	return [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16)) as [number, number, number];
}

const clamp255 = (v: number): number => Math.max(0, Math.min(255, v));

/** `#RRGGBB` (upper case) from rounded, clamped 0-255 channels. */
export const rgbToUpperHex = (rgb: readonly number[]): string =>
	`#${rgb
		.map((v) => clamp255(Math.round(v)).toString(16).padStart(2, '0'))
		.join('')
		.toUpperCase()}`;
