/**
 * JBIG2 pattern dictionaries and halftone regions (ITU-T T.88 6.7, 6.6 and the gray-scale image
 * decoding of Annex C.5): patterns placed on a grid by gray-scale values coded as Gray-coded
 * bitplanes, arithmetic or MMR coded.
 */
import { ArithmeticDecoder } from './jbig2-arith';
import { Bitmap, compose, crop, Jbig2Error, type Jbig2Budget } from './jbig2-bitmap';
import { decodeGeneric, decodeMmr, DEFAULT_AT, genericContexts, type At } from './jbig2-generic';

/** Decodes a pattern dictionary segment's data (7.4.4, 6.7.5). */
export function decodePatternDictionary(
	data: Uint8Array,
	start: number,
	end: number,
	budget: Jbig2Budget,
): Bitmap[] {
	if (end - start < 7) throw new Jbig2Error('A JBIG2 pattern dictionary is too short.');
	const flags = data[start]!;
	const mmr = (flags & 1) === 1;
	const template = (flags >> 1) & 3;
	const width = data[start + 1]!;
	const height = data[start + 2]!;
	const grayMax = new DataView(data.buffer, data.byteOffset + start + 3, 4).getUint32(0);
	if (!width || !height) throw new Jbig2Error('A JBIG2 pattern has no size.');
	const count = grayMax + 1;
	if (count * width > budget.maxPixels / Math.max(1, height))
		throw new Jbig2Error('A JBIG2 pattern dictionary is larger than the size limit.');
	const collectiveWidth = count * width;
	const at: At[] =
		template === 0
			? [
					[-width, 0],
					[-3, -1],
					[2, -2],
					[-2, -2],
				]
			: [[-width, 0]];
	const collective = mmr
		? decodeMmr(data, start + 7, end, collectiveWidth, height, budget).bitmap
		: decodeGeneric(
				new ArithmeticDecoder(data, start + 7, end),
				genericContexts(template),
				{ width: collectiveWidth, height, template, at, tpgdon: false },
				budget,
			);
	const patterns: Bitmap[] = [];
	for (let index = 0; index < count; index++)
		patterns.push(crop(collective, index * width, 0, width, height, budget));
	return patterns;
}

export interface HalftoneRegion {
	width: number;
	height: number;
	flags: number;
	gridWidth: number;
	gridHeight: number;
	gridX: number;
	gridY: number;
	stepX: number;
	stepY: number;
}

/** Reads the halftone region header after the region information (7.4.5.1). */
export function readHalftoneHeader(
	data: Uint8Array,
	at: number,
	end: number,
	width: number,
	height: number,
): HalftoneRegion & { start: number } {
	if (at + 21 > end) throw new Jbig2Error('A JBIG2 halftone region is too short.');
	const view = new DataView(data.buffer, data.byteOffset + at, 21);
	return {
		width,
		height,
		flags: view.getUint8(0),
		gridWidth: view.getUint32(1),
		gridHeight: view.getUint32(5),
		gridX: view.getInt32(9),
		gridY: view.getInt32(13),
		stepX: view.getUint16(17),
		stepY: view.getUint16(19),
		start: at + 21,
	};
}

/** Decodes a halftone region (6.6.5) from `data[start, end)` with the dictionary's patterns. */
export function decodeHalftone(
	region: HalftoneRegion,
	patterns: Bitmap[],
	data: Uint8Array,
	start: number,
	end: number,
	budget: Jbig2Budget,
): Bitmap {
	const { width, height, flags, gridWidth: gw, gridHeight: gh } = region;
	const mmr = (flags & 1) === 1;
	const template = (flags >> 1) & 3;
	const enableSkip = (flags & 8) !== 0;
	const operator = (flags >> 4) & 7;
	const out = new Bitmap(width, height, budget, (flags >> 7) & 1);
	const first = patterns[0];
	if (!first) throw new Jbig2Error('A JBIG2 halftone region has no patterns.');
	if (gw * gh > budget.maxPixels) throw new Jbig2Error('A JBIG2 halftone grid is too large.');
	const place = (m: number, n: number): [number, number] => [
		Math.floor((region.gridX + m * region.stepY + n * region.stepX) / 256),
		Math.floor((region.gridY + m * region.stepX - n * region.stepY) / 256),
	];
	let skip: Bitmap | undefined;
	if (enableSkip) {
		skip = new Bitmap(gw, gh, budget);
		for (let m = 0; m < gh; m++)
			for (let n = 0; n < gw; n++) {
				const [x, y] = place(m, n);
				if (x + first.width <= 0 || x >= width || y + first.height <= 0 || y >= height)
					skip.data[m * gw + n] = 1;
			}
	}
	let bits = 0;
	while (1 << bits < patterns.length) bits++;
	const values = grayScale(data, start, end, { gw, gh, bits, mmr, template, skip }, budget);
	for (let m = 0; m < gh; m++)
		for (let n = 0; n < gw; n++) {
			const [x, y] = place(m, n);
			const pattern = patterns[Math.min(values[m * gw + n]!, patterns.length - 1)]!;
			compose(out, pattern, x, y, operator, budget);
		}
	return out;
}

/** Decodes the gray-scale image (C.5): `bits` Gray-coded bitplanes, most significant first. */
function grayScale(
	data: Uint8Array,
	start: number,
	end: number,
	options: {
		gw: number;
		gh: number;
		bits: number;
		mmr: boolean;
		template: number;
		skip: Bitmap | undefined;
	},
	budget: Jbig2Budget,
): Uint32Array {
	const { gw, gh, bits, mmr, template, skip } = options;
	const values = new Uint32Array(gw * gh);
	const decoder = mmr ? undefined : new ArithmeticDecoder(data, start, end);
	const contexts = genericContexts(template);
	let at = start;
	let previous: Bitmap | undefined;
	for (let plane = bits - 1; plane >= 0; plane--) {
		let bitmap: Bitmap;
		if (decoder) {
			const generic = { width: gw, height: gh, template, at: DEFAULT_AT[template]!, tpgdon: false };
			bitmap = decodeGeneric(decoder, contexts, skip ? { ...generic, skip } : generic, budget);
		} else {
			const decoded = decodeMmr(data, at, end, gw, gh, budget);
			bitmap = decoded.bitmap;
			at = decoded.end;
		}
		const plain = bitmap.data;
		if (previous)
			for (let index = 0; index < plain.length; index++)
				plain[index] = plain[index]! ^ previous.data[index]!;
		for (let index = 0; index < values.length; index++)
			values[index] = values[index]! | (plain[index]! << plane);
		previous = bitmap;
	}
	return values;
}
