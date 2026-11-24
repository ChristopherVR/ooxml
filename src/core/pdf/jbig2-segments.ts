/**
 * JBIG2 segment headers (ITU-T T.88 7.2) in the three organisations: an embedded stream as PDF
 * holds it (7.4.7 of ISO 32000-1, no file header), and the sequential and random-access file
 * organisations (Annex D) behind a file header.
 */
import { Jbig2Error } from './jbig2-bitmap';

export interface Segment {
	number: number;
	type: number;
	referred: number[];
	page: number;
	/** The segment data, `data[start, end)`. */
	start: number;
	end: number;
	/** An immediate generic region whose length was unknown: its data ends with a row count. */
	unknownLength?: true;
}

const FILE_ID = [0x97, 0x4a, 0x42, 0x32, 0x0d, 0x0a, 0x1a, 0x0a];

export const isJbig2File = (data: Uint8Array): boolean =>
	FILE_ID.every((byte, index) => data[index] === byte);

const u32 = (data: Uint8Array, at: number): number => {
	if (at + 4 > data.length) throw new Jbig2Error('The JBIG2 data ends early.');
	return ((data[at]! << 24) | (data[at + 1]! << 16) | (data[at + 2]! << 8) | data[at + 3]!) >>> 0;
};

type Header = Omit<Segment, 'start' | 'end'> & { length: number; size: number };

function readHeader(data: Uint8Array, at: number): Header {
	const begin = at;
	const number = u32(data, at);
	at += 4;
	const flags = data[at++];
	const retention = data[at];
	if (flags === undefined || retention === undefined)
		throw new Jbig2Error('The JBIG2 data ends early.');
	let count = retention >> 5;
	if (count === 7) {
		count = u32(data, at) & 0x1fffffff;
		at += 4 + Math.ceil((count + 1) / 8);
	} else if (count > 4) throw new Jbig2Error('A JBIG2 segment header is invalid.');
	else at++;
	const size = number <= 256 ? 1 : number <= 65_536 ? 2 : 4;
	if (at + count * size > data.length) throw new Jbig2Error('The JBIG2 data ends early.');
	const referred: number[] = [];
	for (let index = 0; index < count; index++) {
		let value = 0;
		for (let byte = 0; byte < size; byte++) value = value * 256 + data[at++]!;
		referred.push(value);
	}
	const page = flags & 64 ? u32(data, at) : (data[at] ?? 0);
	at += flags & 64 ? 4 : 1;
	const length = u32(data, at);
	at += 4;
	return { number, type: flags & 63, referred, page, length, size: at - begin };
}

/** The end of an immediate generic region of unknown length (7.2.7): after its row count. */
function unknownLengthEnd(data: Uint8Array, start: number): number {
	const flags = data[start + 17] ?? 0;
	const mmr = flags & 1;
	const atBytes = mmr ? 0 : flags & 6 ? 2 : 8;
	for (let at = start + 18 + atBytes; at + 6 <= data.length; at++)
		if (mmr ? data[at] === 0 && data[at + 1] === 0 : data[at] === 0xff && data[at + 1] === 0xac)
			return at + 6;
	throw new Jbig2Error('A JBIG2 generic region of unknown length has no end marker.');
}

/** The segments of an embedded stream or a JBIG2 file. */
export function readSegments(data: Uint8Array): Segment[] {
	let at = 0;
	let sequential = true;
	if (isJbig2File(data)) {
		const flags = data[8] ?? 0;
		sequential = (flags & 1) === 1;
		at = 9 + (flags & 2 ? 0 : 4);
	}
	const segments: Segment[] = [];
	if (sequential) {
		while (at < data.length) {
			const header = readHeader(data, at);
			const start = at + header.size;
			let end = start + header.length;
			const unknown = header.length === 0xffffffff;
			if (unknown) {
				if (header.type !== 38)
					throw new Jbig2Error('A JBIG2 segment other than a generic region has no length.');
				end = unknownLengthEnd(data, start);
			}
			if (end > data.length) throw new Jbig2Error('A JBIG2 segment ends after the data.');
			const { length: _length, size: _size, ...rest } = header;
			segments.push(
				unknown ? { ...rest, start, end, unknownLength: true } : { ...rest, start, end },
			);
			at = end;
			if (header.type === 51) break;
		}
		return segments;
	}
	// Random access: every header (up to the end of file segment), then every segment's data.
	const headers: Header[] = [];
	while (at < data.length) {
		const header = readHeader(data, at);
		if (header.length === 0xffffffff)
			throw new Jbig2Error('A JBIG2 segment in a random-access file has no length.');
		headers.push(header);
		at += header.size;
		if (header.type === 51) break;
	}
	for (const { length, size: _size, ...rest } of headers) {
		if (at + length > data.length) throw new Jbig2Error('A JBIG2 segment ends after the data.');
		segments.push({ ...rest, start: at, end: at + length });
		at += length;
	}
	return segments;
}

export interface RegionInfo {
	width: number;
	height: number;
	x: number;
	y: number;
	operator: number;
}

/** The region segment information field (7.4.1), 17 bytes at `at`. */
export function readRegionInfo(data: Uint8Array, at: number, end: number): RegionInfo {
	if (at + 17 > end) throw new Jbig2Error('A JBIG2 region segment is too short.');
	return {
		width: u32(data, at),
		height: u32(data, at + 4),
		x: u32(data, at + 8) | 0,
		y: u32(data, at + 12) | 0,
		operator: data[at + 16]! & 7,
	};
}

export { u32 };
