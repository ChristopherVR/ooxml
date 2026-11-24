/**
 * The JP2 file format (ITU-T T.800 Annex I, and the JPX boxes PDF files use): the signature, the
 * header box with its colour specification, palette, component mapping and channel definitions,
 * and the contiguous codestream. A bare codestream (starting with SOC) is accepted as it is.
 */
import { JpxError } from './errors';

export interface JpxPalette {
	/** Number of entries. */
	entries: number;
	/** Per palette column: its bit depth, signedness and values. */
	columns: { depth: number; signed: boolean; values: Int32Array }[];
}

export interface JpxChannel {
	channel: number;
	/** 0 colour, 1 opacity, 2 premultiplied opacity, 65535 unspecified. */
	type: number;
	/** 0 the whole image, 1..n the colour it belongs to, 65535 none. */
	association: number;
}

export interface JpxHeader {
	/** The enumerated colour space (16 sRGB, 17 greyscale, 18 sYCC, 12 CMYK...). */
	enumerated?: number;
	/** The colour space signature of an embedded ICC profile ('RGB ', 'GRAY', 'CMYK', 'Lab '). */
	icc?: string;
	palette?: JpxPalette;
	/** Component mapping: each channel from a component directly or through a palette column. */
	mapping?: { component: number; palette: boolean; column: number }[];
	channels?: JpxChannel[];
}

const u16 = (data: Uint8Array, at: number) => (data[at]! << 8) | data[at + 1]!;
const u32 = (data: Uint8Array, at: number) => u16(data, at) * 65536 + u16(data, at + 2);
const type = (data: Uint8Array, at: number) =>
	String.fromCharCode(data[at]!, data[at + 1]!, data[at + 2]!, data[at + 3]!);

interface Box {
	type: string;
	start: number;
	end: number;
}

function* boxes(data: Uint8Array, start: number, end: number): Generator<Box> {
	let at = start;
	while (at + 8 <= end) {
		let length = u32(data, at);
		const name = type(data, at + 4);
		let header = 8;
		if (length === 1) {
			if (at + 16 > end) return;
			length = u32(data, at + 8) * 2 ** 32 + u32(data, at + 12);
			header = 16;
		} else if (length === 0) length = end - at;
		if (length < header) throw new JpxError('its JPEG 2000 data has a box of invalid length');
		const boxEnd = Math.min(end, at + length);
		yield { type: name, start: at + header, end: boxEnd };
		at += length;
	}
}

function palette(data: Uint8Array, box: Box): JpxPalette {
	const entries = u16(data, box.start);
	const count = data[box.start + 2]!;
	if (!entries || entries > 1024 || !count) throw new JpxError('its JPEG 2000 palette is invalid');
	const columns = Array.from({ length: count }, (_, index) => {
		const spec = data[box.start + 3 + index]!;
		return {
			depth: (spec & 0x7f) + 1,
			signed: (spec & 0x80) !== 0,
			values: new Int32Array(entries),
		};
	});
	let at = box.start + 3 + count;
	for (let entry = 0; entry < entries; entry++)
		for (const column of columns) {
			const bytes = Math.ceil(column.depth / 8);
			if (at + bytes > box.end) throw new JpxError('its JPEG 2000 palette is truncated');
			let value = 0;
			for (let byte = 0; byte < bytes; byte++) value = value * 256 + data[at++]!;
			if (column.signed && value >= 2 ** (column.depth - 1)) value -= 2 ** column.depth;
			column.values[entry] = value;
		}
	return { entries, columns };
}

function readHeader(data: Uint8Array, box: Box, header: JpxHeader): void {
	for (const child of boxes(data, box.start, box.end)) {
		const at = child.start;
		if (child.type === 'colr' && header.enumerated === undefined && header.icc === undefined) {
			const method = data[at];
			if (method === 1 && child.end - at >= 7) header.enumerated = u32(data, at + 3);
			else if ((method === 2 || method === 3) && child.end - at >= 3 + 20)
				header.icc = type(data, at + 3 + 16);
		} else if (child.type === 'pclr') header.palette = palette(data, child);
		else if (child.type === 'cmap')
			header.mapping = Array.from({ length: Math.floor((child.end - at) / 4) }, (_, index) => ({
				component: u16(data, at + index * 4),
				palette: data[at + index * 4 + 2] === 1,
				column: data[at + index * 4 + 3]!,
			}));
		else if (child.type === 'cdef') {
			const count = Math.min(u16(data, at), Math.floor((child.end - at - 2) / 6));
			header.channels = Array.from({ length: count }, (_, index) => ({
				channel: u16(data, at + 2 + index * 6),
				type: u16(data, at + 4 + index * 6),
				association: u16(data, at + 6 + index * 6),
			}));
		}
	}
}

/** The codestream of a JP2/JPX file or bare codestream, with the header boxes that describe it. */
export function readJp2(data: Uint8Array): { codestream: Uint8Array; header: JpxHeader } {
	if (data[0] === 0xff && data[1] === 0x4f) return { codestream: data, header: {} };
	const header: JpxHeader = {};
	let signature = false;
	for (const box of boxes(data, 0, data.length)) {
		if (box.type === 'jP  ') signature = true;
		else if (box.type === 'jp2h') readHeader(data, box, header);
		else if (box.type === 'jp2c') {
			if (!signature) break;
			return { codestream: data.subarray(box.start, box.end), header };
		}
	}
	throw new JpxError(
		signature ? 'its JPEG 2000 data has no codestream' : 'it is not a JPEG 2000 file',
	);
}
