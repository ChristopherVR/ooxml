/**
 * Stream filters (ISO 32000-1 section 7.4): Flate (with PNG and TIFF predictors), LZW, ASCIIHex,
 * ASCII85 and RunLength. Image codecs (DCT, JPX, JBIG2, CCITT) are left encoded for the caller.
 * Every decoder stops at `maxBytes`, so a small stream cannot expand without bound.
 */
import { Unzlib, Inflate } from 'fflate';
import { hexValue, isWhite } from './lexer';
import { isDict, isName, isNumber, type PdfDict } from './objects';

export class PdfLimitError extends Error {}

/** Filters left to the image reader. */
export const IMAGE_FILTERS = new Set(['DCTDecode', 'JPXDecode', 'JBIG2Decode', 'CCITTFaxDecode']);
const ABBREVIATIONS: Record<string, string> = {
	Fl: 'FlateDecode',
	LZW: 'LZWDecode',
	AHx: 'ASCIIHexDecode',
	A85: 'ASCII85Decode',
	RL: 'RunLengthDecode',
	DCT: 'DCTDecode',
	CCF: 'CCITTFaxDecode',
};

/** The filter names and parameter dictionaries of a stream, in order. */
export function streamFilters(dict: PdfDict): { name: string; parms: PdfDict | undefined }[] {
	const filter = dict.get('Filter') ?? dict.get('F');
	const parms = dict.get('DecodeParms') ?? dict.get('DP');
	const names = Array.isArray(filter) ? filter : filter ? [filter] : [];
	const params = Array.isArray(parms) ? parms : [parms];
	return names.flatMap((name, index) =>
		isName(name)
			? [
					{
						name: ABBREVIATIONS[name.name] ?? name.name,
						parms: isDict(params[index]) ? (params[index] as PdfDict) : undefined,
					},
				]
			: [],
	);
}

class Sink {
	readonly chunks: Uint8Array[] = [];
	size = 0;
	constructor(readonly max: number) {}
	push(chunk: Uint8Array): void {
		this.size += chunk.length;
		if (this.size > this.max)
			throw new PdfLimitError('A PDF stream decodes to more than the size limit.');
		this.chunks.push(chunk);
	}
	bytes(): Uint8Array {
		if (this.chunks.length === 1) return this.chunks[0]!;
		const out = new Uint8Array(this.size);
		let at = 0;
		for (const chunk of this.chunks) {
			out.set(chunk, at);
			at += chunk.length;
		}
		return out;
	}
}

function flate(data: Uint8Array, max: number): Uint8Array {
	const run = (raw: boolean) => {
		const sink = new Sink(max);
		const stream = raw ? new Inflate() : new Unzlib();
		stream.ondata = (chunk) => sink.push(chunk);
		try {
			stream.push(data, true);
		} catch (error) {
			// A truncated stream keeps what decoded before the damage, as viewers do.
			if (error instanceof PdfLimitError || !sink.size) throw error;
		}
		return sink.bytes();
	};
	try {
		return run(false);
	} catch (error) {
		if (error instanceof PdfLimitError) throw error;
		return run(true);
	}
}

function lzw(data: Uint8Array, max: number, earlyChange: number): Uint8Array {
	const sink: number[] = [];
	let table: number[][] = [];
	const reset = () => {
		table = Array.from({ length: 258 }, (_, index) => (index < 256 ? [index] : []));
	};
	reset();
	let codeLength = 9;
	let buffer = 0;
	let bits = 0;
	let previous: number[] | undefined;
	for (const byte of data) {
		buffer = (buffer << 8) | byte;
		bits += 8;
		while (bits >= codeLength) {
			const code = (buffer >> (bits - codeLength)) & ((1 << codeLength) - 1);
			bits -= codeLength;
			if (code === 256) {
				reset();
				codeLength = 9;
				previous = undefined;
				continue;
			}
			if (code === 257) return Uint8Array.from(sink);
			let entry = table[code];
			if (!entry || (!entry.length && code >= 258)) {
				if (!previous) return Uint8Array.from(sink);
				entry = [...previous, previous[0]!];
			}
			sink.push(...entry);
			if (sink.length > max)
				throw new PdfLimitError('A PDF stream decodes to more than the size limit.');
			if (previous) table.push([...previous, entry[0]!]);
			previous = entry;
			if (table.length + earlyChange >= 1 << codeLength && codeLength < 12) codeLength++;
		}
	}
	return Uint8Array.from(sink);
}

function asciiHex(data: Uint8Array): Uint8Array {
	const out: number[] = [];
	let high = -1;
	for (const byte of data) {
		if (byte === 0x3e) break;
		const digit = hexValue(byte);
		if (digit < 0) continue;
		if (high < 0) high = digit;
		else {
			out.push((high << 4) | digit);
			high = -1;
		}
	}
	if (high >= 0) out.push(high << 4);
	return Uint8Array.from(out);
}

function ascii85(data: Uint8Array, max: number): Uint8Array {
	const out: number[] = [];
	const group: number[] = [];
	const flush = (count: number) => {
		while (group.length < 5) group.push(84);
		let value = 0;
		for (const digit of group) value = value * 85 + digit;
		for (let index = 0; index < count - 1; index++) out.push((value >>> (24 - index * 8)) & 0xff);
		group.length = 0;
	};
	let at = data[0] === 0x3c && data[1] === 0x7e ? 2 : 0;
	for (; at < data.length; at++) {
		const byte = data[at]!;
		if (byte === 0x7e) break;
		if (isWhite(byte)) continue;
		if (byte === 0x7a && group.length === 0) {
			out.push(0, 0, 0, 0);
			if (out.length > max)
				throw new PdfLimitError('A PDF stream decodes to more than the size limit.');
			continue;
		}
		if (byte < 0x21 || byte > 0x75) continue;
		group.push(byte - 0x21);
		if (group.length === 5) flush(5);
		if (out.length > max)
			throw new PdfLimitError('A PDF stream decodes to more than the size limit.');
	}
	if (group.length) flush(group.length);
	return Uint8Array.from(out);
}

function runLength(data: Uint8Array, max: number): Uint8Array {
	const out: number[] = [];
	for (let at = 0; at < data.length;) {
		const length = data[at++]!;
		if (length === 128) break;
		if (length < 128) {
			for (let index = 0; index <= length && at < data.length; index++) out.push(data[at++]!);
		} else {
			const byte = data[at++] ?? 0;
			for (let index = 0; index < 257 - length; index++) out.push(byte);
		}
		if (out.length > max)
			throw new PdfLimitError('A PDF stream decodes to more than the size limit.');
	}
	return Uint8Array.from(out);
}

const parm = (parms: PdfDict | undefined, key: string, fallback: number): number => {
	const value = parms?.get(key);
	return isNumber(value) ? value : fallback;
};

/** Undoes a PNG (10..15) or TIFF (2) predictor. */
export function unpredict(data: Uint8Array, parms: PdfDict | undefined): Uint8Array {
	const predictor = parm(parms, 'Predictor', 1);
	if (predictor <= 1) return data;
	const colors = parm(parms, 'Colors', 1);
	const bpc = parm(parms, 'BitsPerComponent', 8);
	const columns = parm(parms, 'Columns', 1);
	const bpp = Math.max(1, Math.ceil((colors * bpc) / 8));
	const stride = Math.ceil((colors * bpc * columns) / 8);
	if (predictor === 2) {
		if (bpc !== 8) return data;
		const out = data.slice();
		for (let row = 0; row + stride <= out.length; row += stride)
			for (let at = row + bpp; at < row + stride; at++)
				out[at] = (out[at]! + out[at - bpp]!) & 0xff;
		return out;
	}
	const rows = Math.floor(data.length / (stride + 1));
	const out = new Uint8Array(rows * stride);
	for (let row = 0; row < rows; row++) {
		const type = data[row * (stride + 1)]!;
		const source = row * (stride + 1) + 1;
		const target = row * stride;
		for (let x = 0; x < stride; x++) {
			const left = x >= bpp ? out[target + x - bpp]! : 0;
			const up = row ? out[target - stride + x]! : 0;
			const corner = row && x >= bpp ? out[target - stride + x - bpp]! : 0;
			const value = data[source + x]!;
			let predicted = 0;
			if (type === 1) predicted = left;
			else if (type === 2) predicted = up;
			else if (type === 3) predicted = (left + up) >> 1;
			else if (type === 4) {
				const p = left + up - corner;
				const pa = Math.abs(p - left);
				const pb = Math.abs(p - up);
				const pc = Math.abs(p - corner);
				predicted = pa <= pb && pa <= pc ? left : pb <= pc ? up : corner;
			}
			out[target + x] = (value + predicted) & 0xff;
		}
	}
	return out;
}

/**
 * Decodes the stream filters before the first image codec. Returns the bytes and the image
 * filter left, if any (the bytes are then that codec's input).
 */
export function decodeFilters(
	data: Uint8Array,
	dict: PdfDict,
	max: number,
): { bytes: Uint8Array; imageFilter?: { name: string; parms: PdfDict | undefined } } {
	let bytes = data;
	for (const filter of streamFilters(dict)) {
		if (IMAGE_FILTERS.has(filter.name)) return { bytes, imageFilter: filter };
		if (filter.name === 'FlateDecode') bytes = unpredict(flate(bytes, max), filter.parms);
		else if (filter.name === 'LZWDecode')
			bytes = unpredict(lzw(bytes, max, parm(filter.parms, 'EarlyChange', 1)), filter.parms);
		else if (filter.name === 'ASCIIHexDecode') bytes = asciiHex(bytes);
		else if (filter.name === 'ASCII85Decode') bytes = ascii85(bytes, max);
		else if (filter.name === 'RunLengthDecode') bytes = runLength(bytes, max);
		else if (filter.name !== 'Crypt') throw new Error(`Unsupported PDF filter ${filter.name}.`);
	}
	return { bytes };
}
