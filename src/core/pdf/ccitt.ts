/**
 * The CCITTFaxDecode filter (ISO 32000-1 section 7.4.6): ITU-T T.4 Group 3 one-dimensional
 * (Modified Huffman, K = 0) and mixed one- and two-dimensional (K > 0) coding, and T.6 Group 4
 * (K < 0). Decodes into 1-bit rows (0 is black unless BlackIs1), stopping at the row count, the
 * end of the data, an end-of-block code, a damaged row or `maxBytes`, whichever comes first.
 */
import {
	BLACK_CODES,
	EOL_CODE,
	MAX_CODE_BITS,
	MODE_CODES,
	MODE_HORIZONTAL,
	MODE_PASS,
	MODE_VERTICAL,
	WHITE_CODES,
	type CodeTable,
} from './ccitt-codes';

export interface CcittOptions {
	/** < 0: Group 4; 0: Group 3 one-dimensional; > 0: Group 3 two-dimensional. */
	k: number;
	columns: number;
	/** Rows to decode (the image height). */
	rows: number;
	encodedByteAlign?: boolean;
	blackIs1?: boolean;
}

class Bits {
	at = 0;
	constructor(readonly data: Uint8Array) {}
	get done(): boolean {
		return this.at >= this.data.length * 8;
	}
	bit(offset = 0): number {
		const at = this.at + offset;
		const byte = this.data[at >> 3];
		return byte === undefined ? 0 : (byte >> (7 - (at & 7))) & 1;
	}
	align(): void {
		this.at = (this.at + 7) & ~7;
	}
	/** Fill zeros and an end-of-line code (eleven or more zeros, then a one) at `offset`. */
	eolLength(offset = 0): number {
		let zeros = 0;
		while (zeros < 64 && !this.done && this.bit(offset + zeros) === 0) zeros++;
		return zeros >= 11 &&
			this.at + offset + zeros < this.data.length * 8 &&
			this.bit(offset + zeros)
			? zeros + 1
			: 0;
	}
	code(table: CodeTable): number | undefined {
		let code = 0;
		for (let length = 1; length <= MAX_CODE_BITS; length++) {
			code = (code << 1) | this.bit();
			this.at++;
			const value = table.get((1 << length) | code);
			if (value !== undefined) return value;
		}
		return undefined;
	}
	/** A run length (make-up codes, then a terminating code); -1 when damaged. */
	run(color: number): number {
		let total = 0;
		for (let guard = 0; guard < 64; guard++) {
			const value = this.code(color ? BLACK_CODES : WHITE_CODES);
			if (value === undefined || value === EOL_CODE) return -1;
			total += value;
			if (value < 64) return total;
		}
		return -1;
	}
}

/** One-dimensional row: the positions where the colour changes; undefined when damaged. */
function row1d(bits: Bits, columns: number): number[] | undefined {
	const changes: number[] = [];
	let a0 = 0;
	let color = 0;
	while (a0 < columns) {
		const run = bits.run(color);
		if (run < 0) return undefined;
		a0 = Math.min(columns, a0 + run);
		changes.push(a0);
		color ^= 1;
	}
	return changes;
}

/** Two-dimensional row against the reference row's changes (ending with two `columns`). */
function row2d(bits: Bits, columns: number, reference: number[]): number[] | undefined {
	const changes: number[] = [];
	let a0 = -1;
	let color = 0;
	let j = 0;
	for (let guard = 0; a0 < columns; guard++) {
		if (guard > columns * 2 + 8) return undefined;
		while (j > 0 && reference[j - 1]! > a0) j--;
		while (j < reference.length - 1 && (reference[j]! <= a0 || (j & 1) !== color)) j++;
		const b1 = reference[j]!;
		const b2 = reference[j + 1] ?? columns;
		const mode = bits.code(MODE_CODES);
		if (mode === MODE_PASS) a0 = b2;
		else if (mode === MODE_HORIZONTAL) {
			const start = Math.max(a0, 0);
			const first = bits.run(color);
			const second = bits.run(color ^ 1);
			if (first < 0 || second < 0) return undefined;
			const a1 = Math.min(columns, start + first);
			const a2 = Math.min(columns, a1 + second);
			changes.push(a1, a2);
			a0 = a2;
		} else if (mode !== undefined && Math.abs(mode - MODE_VERTICAL) <= 3) {
			const a1 = Math.min(columns, Math.max(Math.max(a0, 0), b1 + mode - MODE_VERTICAL));
			changes.push(a1);
			a0 = a1;
			color ^= 1;
		} else return undefined;
	}
	return changes;
}

/** Decodes CCITT fax data into packed 1-bit rows. */
export function decodeCcitt(data: Uint8Array, options: CcittOptions, maxBytes: number): Uint8Array {
	return decodeCcittBits(data, options, maxBytes).bytes;
}

/** Decodes CCITT fax data; also returns the bit position where decoding stopped. */
export function decodeCcittBits(
	data: Uint8Array,
	options: CcittOptions,
	maxBytes: number,
): { bytes: Uint8Array; bits: number } {
	const columns = Math.floor(options.columns);
	if (columns < 1 || columns > 65_536) throw new Error('The CCITT image width is out of range.');
	const stride = (columns + 7) >> 3;
	const rows = Math.max(0, Math.min(Math.floor(options.rows), Math.floor(maxBytes / stride)));
	const white = options.blackIs1 ? 0 : 0xff;
	const out = new Uint8Array(stride * rows).fill(white);
	const bits = new Bits(data);
	let reference = [columns, columns];
	for (let row = 0; row < rows; row++) {
		if (options.encodedByteAlign && options.k < 0) bits.align();
		let eols = 0;
		for (;;) {
			const length = bits.eolLength();
			if (length) {
				bits.at += length;
				eols++;
				continue;
			}
			// Group 3 two-dimensional: a tag bit follows each end of line (RTC is EOL 1 EOL 1...).
			if (options.k > 0 && eols && bits.bit() === 1 && bits.eolLength(1)) {
				bits.at++;
				continue;
			}
			break;
		}
		if (eols >= 2 || bits.done) break;
		let twoD = options.k < 0;
		if (options.k > 0) {
			twoD = bits.bit() === 0;
			bits.at++;
		}
		const changes = twoD ? row2d(bits, columns, reference) : row1d(bits, columns);
		if (!changes) break;
		let color = 0;
		let from = 0;
		for (const change of [...changes, columns]) {
			if (color && change > from) fillBlack(out, row * stride, from, change, white === 0);
			from = Math.max(from, change);
			color ^= 1;
		}
		reference = [...changes.filter((change) => change < columns), columns, columns];
		if (options.encodedByteAlign && options.k >= 0) bits.align();
	}
	return { bytes: out, bits: bits.at };
}

function fillBlack(out: Uint8Array, start: number, from: number, to: number, set: boolean): void {
	for (let x = from; x < to; x++) {
		const index = start + (x >> 3);
		const mask = 0x80 >> (x & 7);
		out[index] = set ? out[index]! | mask : out[index]! & ~mask;
	}
}
