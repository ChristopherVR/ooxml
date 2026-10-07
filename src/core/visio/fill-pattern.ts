import { encodePng } from '@christophervr/ole2/utils/png-encoder';
import { clampUnitInterval, hexToRgbChannels } from '../color/color-primitives';
import { FILL_PATTERN_TILES } from './fill-pattern-tiles';
import { number, type Cells } from './sheet';
import type { VisioFillPattern } from './model';

/** Native GDI byte alpha, premultiplication and tile-coverage quantization. */
export function fillPatternPixels(
	pattern: number,
	foreground: string,
	background: string,
	foregroundOpacity: number,
	backgroundOpacity: number,
): Uint8Array | undefined {
	const rows = FILL_PATTERN_TILES[pattern];
	const front = hexToRgbChannels(foreground),
		back = hexToRgbChannels(background);
	if (!rows || !front || !back) return undefined;
	const fa = Math.floor(255 * clampUnitInterval(foregroundOpacity));
	const ba = Math.floor(255 * clampUnitInterval(backgroundOpacity));
	const fc = [front.r, front.g, front.b].map((value) => Math.round((value * fa) / 255));
	const bc = [back.r, back.g, back.b].map((value) => Math.round((value * ba) / 255));
	const pixels = new Uint8Array(8 * 8 * 4);
	for (let y = 0; y < 8; y++)
		for (let x = 0; x < 8; x++) {
			const weight = [0, 1, 0.25, 234 / 256][Number(rows[y]![x])]!;
			const alpha = Math.floor(fa * weight + ba * (1 - weight));
			const offset = (y * 8 + x) * 4;
			for (let channel = 0; channel < 3; channel++)
				pixels[offset + channel] = alpha
					? Math.floor(
							(Math.floor(fc[channel]! * weight + bc[channel]! * (1 - weight)) * 255) / alpha,
						)
					: 0;
			pixels[offset + 3] = alpha;
		}
	return pixels;
}

export function cachedFillPattern(
	cells: Cells,
	resolve: (name: string) => string,
): VisioFillPattern | undefined {
	if (!FILL_PATTERN_TILES[number(cells, 'FillPattern', 1)]) return undefined;
	const pixels = fillPatternPixels(
		number(cells, 'FillPattern', 1),
		resolve('FillForegnd'),
		resolve('FillBkgnd'),
		1 - number(cells, 'FillForegndTrans', 0),
		1 - number(cells, 'FillBkgndTrans', 0),
	);
	return pixels
		? {
				width: 1 / 12,
				height: 1 / 12,
				mimeType: 'image/png',
				pixelWidth: 8,
				pixelHeight: 8,
				bytes: encodePng(8, 8, pixels),
			}
		: undefined;
}
