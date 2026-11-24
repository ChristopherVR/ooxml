/**
 * JBIG2 Huffman decoding (ITU-T T.88 Annex B): the MSB-first bit reader, prefix code assignment
 * (B.3), the standard tables B.1 to B.15 and user tables from table segments (B.2, type 53).
 */
import { Jbig2Error } from './jbig2-bitmap';
import { standardLines, type TableLine } from './jbig2-tables';

export type { TableLine } from './jbig2-tables';

export class BitReader {
	/** Position in bits from the start of `data`. */
	at: number;
	constructor(
		readonly data: Uint8Array,
		start: number,
		readonly end: number,
	) {
		this.at = start * 8;
	}
	bit(): number {
		const byte = this.at >> 3;
		if (byte >= this.end) throw new Jbig2Error('The JBIG2 Huffman data ends early.');
		const value = (this.data[byte]! >> (7 - (this.at & 7))) & 1;
		this.at++;
		return value;
	}
	bits(count: number): number {
		let value = 0;
		for (let index = 0; index < count; index++) value = value * 2 + this.bit();
		return value;
	}
	align(): void {
		this.at = (this.at + 7) & ~7;
	}
	/** The byte offset (after aligning). */
	get byte(): number {
		return (this.at + 7) >> 3;
	}
}

export interface HuffmanTable {
	/** (1 << length | code) to the line it selects. */
	codes: Map<number, TableLine>;
	maxLength: number;
}

/** Assigns prefix codes (B.3) to the lines, in their order. */
export function buildTable(lines: TableLine[]): HuffmanTable {
	const maxLength = lines.reduce((max, line) => Math.max(max, line.prefix), 0);
	if (maxLength > 32) throw new Jbig2Error('A JBIG2 Huffman code is too long.');
	const counts = new Array<number>(maxLength + 1).fill(0);
	for (const line of lines) counts[line.prefix]!++;
	counts[0] = 0;
	const codes = new Map<number, TableLine>();
	let first = 0;
	for (let length = 1; length <= maxLength; length++) {
		first = (first + counts[length - 1]!) * 2;
		let code = first;
		for (const line of lines)
			if (line.prefix === length) {
				if (code >= 2 ** length)
					throw new Jbig2Error('A JBIG2 Huffman table is not a prefix code.');
				codes.set(2 ** length + code++, line);
			}
	}
	return { codes, maxLength };
}

/** Decodes one value; null is OOB. */
export function decodeHuffman(reader: BitReader, table: HuffmanTable): number | null {
	let code = 0;
	for (let length = 1; length <= table.maxLength; length++) {
		code = code * 2 + reader.bit();
		const line = table.codes.get(2 ** length + code);
		if (!line) continue;
		if (line.kind === 'oob') return null;
		const offset = reader.bits(line.range);
		return line.kind === 'lower' ? line.low - offset : line.low + offset;
	}
	throw new Jbig2Error('The JBIG2 data has an invalid Huffman code.');
}

const standardCache = new Map<number, HuffmanTable>();

/** Standard table B.`number`. */
export function standardTable(number: number): HuffmanTable {
	let table = standardCache.get(number);
	if (!table) standardCache.set(number, (table = buildTable(standardLines(number))));
	return table;
}

/** A user table from a table segment's data (B.2). */
export function readTableSegment(data: Uint8Array, start: number, end: number): HuffmanTable {
	if (end - start < 9) throw new Jbig2Error('A JBIG2 table segment is too short.');
	const flags = data[start]!;
	const view = new DataView(data.buffer, data.byteOffset + start + 1, 8);
	const low = view.getInt32(0);
	const high = view.getInt32(4);
	const prefixBits = ((flags >> 1) & 7) + 1;
	const rangeBits = ((flags >> 4) & 7) + 1;
	const reader = new BitReader(data, start + 9, end);
	const lines: TableLine[] = [];
	for (let current = low; current < high;) {
		const prefix = reader.bits(prefixBits);
		const range = reader.bits(rangeBits);
		lines.push({ prefix, range, low: current });
		current += 2 ** range;
	}
	lines.push({ prefix: reader.bits(prefixBits), range: 32, low: low - 1, kind: 'lower' });
	lines.push({ prefix: reader.bits(prefixBits), range: 32, low: high, kind: 'upper' });
	if (flags & 1) lines.push({ prefix: reader.bits(prefixBits), range: 0, low: 0, kind: 'oob' });
	return buildTable(lines);
}
