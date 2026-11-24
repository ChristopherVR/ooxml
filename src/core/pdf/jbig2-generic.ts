/**
 * JBIG2 generic region decoding (ITU-T T.88 6.2: templates 0 to 3 with adaptive pixels, typical
 * prediction TPGDON, skip bitmaps, and MMR through the CCITT Group 4 decoder) and generic
 * refinement region decoding (6.3: templates 0 and 1, TPGRON).
 */
import { decodeCcittBits } from './ccitt';
import { ArithmeticDecoder } from './jbig2-arith';
import { Bitmap, Jbig2Error, type Jbig2Budget } from './jbig2-bitmap';

/** An adaptive template pixel offset. */
export type At = [number, number];

/** Default AT pixels per template (6.2.5.3 figures; the halftone gray-scale defaults, C.5). */
export const DEFAULT_AT: At[][] = [
	[
		[3, -1],
		[-3, -1],
		[2, -2],
		[-2, -2],
	],
	[[3, -1]],
	[[2, -1]],
	[[2, -1]],
];

/** Context size for each template. */
const CONTEXT_BITS = [16, 13, 10, 10];
/** The SLTP context of typical prediction (6.2.5.7). */
const SLTP = [0x9b25, 0x0795, 0x00e5, 0x0195];

export const genericContexts = (template: number): Uint8Array =>
	new Uint8Array(1 << CONTEXT_BITS[template]!);

/**
 * The template pixels from the most to the least significant context bit, as [dx, dy] offsets;
 * `'a0'`..`'a3'` stand for the adaptive pixels (raster order with their nominal positions).
 */
const TEMPLATES: (At | number)[][] = [
	[3, [-1, -2], [0, -2], [1, -2], 2, 1, [-2, -1], [-1, -1], [0, -1], [1, -1], [2, -1], 0],
	[[-1, -2], [0, -2], [1, -2], [2, -2], [-2, -1], [-1, -1], [0, -1], [1, -1], [2, -1], 0],
	[[-1, -2], [0, -2], [1, -2], [-2, -1], [-1, -1], [0, -1], [1, -1], 0],
	[[-3, -1], [-2, -1], [-1, -1], [0, -1], [1, -1], 0],
];
/** The current-row pixels after the template above, nearest last. */
const ROW_PIXELS = [4, 3, 2, 4];

export interface GenericOptions {
	width: number;
	height: number;
	template: number;
	at: At[];
	tpgdon: boolean;
	skip?: Bitmap;
}

/** Decodes a generic region with the arithmetic decoder (6.2.5). */
export function decodeGeneric(
	decoder: ArithmeticDecoder,
	contexts: Uint8Array,
	options: GenericOptions,
	budget: Jbig2Budget,
): Bitmap {
	const { width, height, template, tpgdon, skip } = options;
	const bitmap = new Bitmap(width, height, budget);
	budget.spend(width * height);
	const offsets: At[] = [];
	for (const entry of TEMPLATES[template]!)
		offsets.push(typeof entry === 'number' ? options.at[entry]! : entry);
	const rowPixels = ROW_PIXELS[template]!;
	const data = bitmap.data;
	let ltp = 0;
	for (let y = 0; y < height; y++) {
		if (tpgdon) {
			ltp ^= decoder.bit(contexts, SLTP[template]!);
			if (ltp) {
				if (y > 0) data.copyWithin(y * width, (y - 1) * width, y * width);
				continue;
			}
		}
		const row = y * width;
		for (let x = 0; x < width; x++) {
			if (skip && skip.data[row + x]) continue;
			let context = 0;
			for (const [dx, dy] of offsets) {
				const px = x + dx;
				const py = y + dy;
				context = (context << 1) | (px >= 0 && px < width && py >= 0 ? data[py * width + px]! : 0);
			}
			for (let dx = rowPixels; dx >= 1; dx--)
				context = (context << 1) | (x >= dx ? data[row + x - dx]! : 0);
			if (decoder.bit(contexts, context)) data[row + x] = 1;
		}
	}
	return bitmap;
}

/**
 * Decodes an MMR (T.6) coded generic region from `data[start, end)`; returns the bitmap and the
 * byte after its data (and the end-of-facsimile-block code, when there is one).
 */
export function decodeMmr(
	data: Uint8Array,
	start: number,
	end: number,
	width: number,
	height: number,
	budget: Jbig2Budget,
): { bitmap: Bitmap; end: number } {
	const bitmap = new Bitmap(width, height, budget);
	budget.spend(width * height);
	if (!width || !height) return { bitmap, end: start };
	if (width > 65_536) throw new Jbig2Error('The JBIG2 MMR region is too wide.');
	const source = data.subarray(start, end);
	const stride = (width + 7) >> 3;
	const decoded = decodeCcittBits(
		source,
		{ k: -1, columns: width, rows: height, blackIs1: true },
		stride * height,
	);
	for (let y = 0; y < height; y++)
		for (let x = 0; x < width; x++)
			if ((decoded.bytes[y * stride + (x >> 3)]! >> (7 - (x & 7))) & 1)
				bitmap.data[y * width + x] = 1;
	let bit = decoded.bits;
	// An EOFB (two EOL codes, 000000000001 twice) may follow the last row.
	const peek = (at: number) => ((source[at >> 3] ?? 0) >> (7 - (at & 7))) & 1;
	const eol = (at: number) => {
		for (let index = 0; index < 11; index++) if (peek(at + index)) return false;
		return peek(at + 11) === 1;
	};
	if (bit + 24 <= source.length * 8 && eol(bit) && eol(bit + 12)) bit += 24;
	return { bitmap, end: Math.min(end, start + ((bit + 7) >> 3)) };
}

export interface RefinementOptions {
	width: number;
	height: number;
	template: number;
	reference: Bitmap;
	dx: number;
	dy: number;
	at: At[];
	tpgron: boolean;
}

export const refinementContexts = (template: number): Uint8Array =>
	new Uint8Array(template ? 1 << 10 : 1 << 13);

/** Decodes a generic refinement region (6.3.5). */
export function decodeRefinement(
	decoder: ArithmeticDecoder,
	contexts: Uint8Array,
	options: RefinementOptions,
	budget: Jbig2Budget,
): Bitmap {
	const { width, height, template, reference: ref, dx, dy, at, tpgron } = options;
	const bitmap = new Bitmap(width, height, budget);
	budget.spend(width * height);
	const get = (x: number, y: number) => bitmap.get(x, y);
	const r = (x: number, y: number) => ref.get(x - dx, y - dy);
	const a1: At = at[0] ?? [-1, -1];
	const a2: At = at[1] ?? [-1, -1];
	const context = template
		? (x: number, y: number) =>
				(r(x, y - 1) << 9) |
				(r(x - 1, y) << 8) |
				(r(x, y) << 7) |
				(r(x + 1, y) << 6) |
				(r(x, y + 1) << 5) |
				(r(x + 1, y + 1) << 4) |
				(get(x - 1, y - 1) << 3) |
				(get(x, y - 1) << 2) |
				(get(x + 1, y - 1) << 1) |
				get(x - 1, y)
		: (x: number, y: number) =>
				(r(x + a2[0], y + a2[1]) << 12) |
				(r(x, y - 1) << 11) |
				(r(x + 1, y - 1) << 10) |
				(r(x - 1, y) << 9) |
				(r(x, y) << 8) |
				(r(x + 1, y) << 7) |
				(r(x - 1, y + 1) << 6) |
				(r(x, y + 1) << 5) |
				(r(x + 1, y + 1) << 4) |
				(get(x + a1[0], y + a1[1]) << 3) |
				(get(x, y - 1) << 2) |
				(get(x + 1, y - 1) << 1) |
				get(x - 1, y);
	let ltp = 0;
	for (let y = 0; y < height; y++) {
		// The SLTP context is the one where only the reference pixel at (x, y) is set (6.3.5.6).
		if (tpgron) ltp ^= decoder.bit(contexts, template ? 1 << 7 : 1 << 8);
		for (let x = 0; x < width; x++) {
			if (ltp) {
				// Typical prediction: a pixel whose 3 x 3 reference neighbourhood is uniform copies it.
				const value = r(x, y);
				let uniform = true;
				for (let j = -1; j <= 1 && uniform; j++)
					for (let i = -1; i <= 1; i++)
						if (r(x + i, y + j) !== value) {
							uniform = false;
							break;
						}
				if (uniform) {
					if (value) bitmap.data[y * width + x] = 1;
					continue;
				}
			}
			if (decoder.bit(contexts, context(x, y))) bitmap.data[y * width + x] = 1;
		}
	}
	return bitmap;
}
