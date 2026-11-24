/**
 * The PDF tokenizer and object parser (ISO 32000-1 sections 7.2 and 7.3), shared by the file
 * structure (objects, cross-reference streams) and content streams. Tolerant of what real files
 * contain: missing whitespace, stray bytes, unbalanced delimiters. Never recursive without a
 * depth limit.
 */
import { PdfName, PdfRef, PdfString, type PdfDict, type PdfValue } from './objects';

const WHITE = new Set([0, 9, 10, 12, 13, 32]);
const DELIMITER = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25]);
const MAX_DEPTH = 100;

export const isWhite = (byte: number | undefined): boolean => byte !== undefined && WHITE.has(byte);
const isRegular = (byte: number | undefined): boolean =>
	byte !== undefined && !WHITE.has(byte) && !DELIMITER.has(byte);

/** A keyword (operator, `obj`, `R`, `true`...). */
export class PdfKeyword {
	constructor(readonly word: string) {}
}

type Token = PdfValue | PdfKeyword | '[' | ']' | '<<' | '>>' | '{' | '}';

/** Marks the end of the input. */
export const EOF = Symbol('eof');

export class PdfLexer {
	pos: number;
	constructor(
		readonly bytes: Uint8Array,
		start = 0,
	) {
		this.pos = start;
	}

	skipSpace(): void {
		const { bytes } = this;
		while (this.pos < bytes.length) {
			const byte = bytes[this.pos]!;
			if (WHITE.has(byte)) this.pos++;
			else if (byte === 0x25) {
				while (this.pos < bytes.length && bytes[this.pos] !== 10 && bytes[this.pos] !== 13)
					this.pos++;
			} else break;
		}
	}

	/** The next token, or `EOF`. */
	token(): Token | typeof EOF {
		this.skipSpace();
		const { bytes } = this;
		if (this.pos >= bytes.length) return EOF;
		const byte = bytes[this.pos]!;
		if (byte === 0x5b || byte === 0x5d || byte === 0x7b || byte === 0x7d) {
			this.pos++;
			return String.fromCharCode(byte) as '[' | ']' | '{' | '}';
		}
		if (byte === 0x3c) {
			if (bytes[this.pos + 1] === 0x3c) {
				this.pos += 2;
				return '<<';
			}
			return this.#hexString();
		}
		if (byte === 0x3e) {
			this.pos += bytes[this.pos + 1] === 0x3e ? 2 : 1;
			return '>>';
		}
		if (byte === 0x28) return this.#literalString();
		if (byte === 0x2f) return this.#name();
		if (byte === 0x29) {
			this.pos++;
			return this.token();
		}
		const start = this.pos;
		while (this.pos < bytes.length && isRegular(bytes[this.pos])) this.pos++;
		const word = latin1(bytes, start, this.pos);
		if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) return Number(word);
		if (/^[+-]?\d*\.?\d*$/.test(word) && /\d/.test(word)) return Number.parseFloat(word) || 0;
		if (word === 'true') return true;
		if (word === 'false') return false;
		if (word === 'null') return null;
		return new PdfKeyword(word);
	}

	#name(): PdfName {
		const { bytes } = this;
		this.pos++;
		const out: number[] = [];
		while (this.pos < bytes.length && isRegular(bytes[this.pos])) {
			const byte = bytes[this.pos]!;
			if (byte === 0x23 && /^[0-9a-f]{2}$/i.test(latin1(bytes, this.pos + 1, this.pos + 3))) {
				out.push(Number.parseInt(latin1(bytes, this.pos + 1, this.pos + 3), 16));
				this.pos += 3;
			} else {
				out.push(byte);
				this.pos++;
			}
		}
		return new PdfName(new TextDecoder('utf-8').decode(Uint8Array.from(out)));
	}

	#hexString(): PdfString {
		const { bytes } = this;
		this.pos++;
		const out: number[] = [];
		let high = -1;
		while (this.pos < bytes.length && bytes[this.pos] !== 0x3e) {
			const digit = hexValue(bytes[this.pos]!);
			this.pos++;
			if (digit < 0) continue;
			if (high < 0) high = digit;
			else {
				out.push((high << 4) | digit);
				high = -1;
			}
		}
		if (high >= 0) out.push(high << 4);
		this.pos++;
		return new PdfString(Uint8Array.from(out));
	}

	#literalString(): PdfString {
		const { bytes } = this;
		this.pos++;
		const out: number[] = [];
		let depth = 1;
		while (this.pos < bytes.length) {
			const byte = bytes[this.pos++]!;
			if (byte === 0x28) depth++;
			else if (byte === 0x29 && --depth === 0) break;
			if (byte !== 0x5c) {
				// An end-of-line in a string is a line feed, whatever bytes wrote it.
				if (byte === 13) {
					if (bytes[this.pos] === 10) this.pos++;
					out.push(10);
				} else out.push(byte);
				continue;
			}
			const next = bytes[this.pos++];
			if (next === undefined) break;
			const escaped: Record<number, number> = {
				0x6e: 10,
				0x72: 13,
				0x74: 9,
				0x62: 8,
				0x66: 12,
				0x28: 0x28,
				0x29: 0x29,
				0x5c: 0x5c,
			};
			if (escaped[next] !== undefined) out.push(escaped[next]!);
			else if (next >= 0x30 && next <= 0x37) {
				let value = next - 0x30;
				for (let count = 0; count < 2; count++) {
					const digit = bytes[this.pos];
					if (digit === undefined || digit < 0x30 || digit > 0x37) break;
					value = value * 8 + (digit - 0x30);
					this.pos++;
				}
				out.push(value & 0xff);
			} else if (next === 13) {
				if (bytes[this.pos] === 10) this.pos++;
			} else if (next !== 10) out.push(next);
		}
		return new PdfString(Uint8Array.from(out));
	}

	/**
	 * The next object: a value, with `n g R` read as a reference. A keyword is returned as is (an
	 * operator in a content stream, `obj` or `stream` in a file).
	 */
	object(depth = 0): PdfValue | PdfKeyword | typeof EOF {
		const token = this.token();
		return this.#build(token, depth);
	}

	#build(token: Token | typeof EOF, depth: number): PdfValue | PdfKeyword | typeof EOF {
		if (depth > MAX_DEPTH) throw new Error('PDF object nesting is too deep.');
		if (token === '[') {
			const items: PdfValue[] = [];
			for (;;) {
				const next = this.token();
				if (next === ']' || next === EOF) return items;
				const value = this.#build(next, depth + 1);
				if (value === EOF) return items;
				if (value instanceof PdfKeyword) {
					if (value.word === 'R' && items.length >= 2) {
						const gen = items.pop();
						const num = items.pop();
						if (typeof num === 'number' && typeof gen === 'number')
							items.push(new PdfRef(num, gen));
					}
					continue;
				}
				items.push(value);
			}
		}
		if (token === '<<') {
			const dict: PdfDict = new Map();
			for (;;) {
				const key = this.token();
				if (key === '>>' || key === EOF) return dict;
				if (!(key instanceof PdfName)) continue;
				const value = this.#value(depth + 1);
				if (value !== undefined) dict.set(key.name, value);
			}
		}
		if (token === ']' || token === '>>' || token === '{' || token === '}')
			return this.object(depth);
		if (typeof token === 'number') {
			// `n g R` is a reference: look ahead without consuming anything else.
			const save = this.pos;
			const gen = this.token();
			if (typeof gen === 'number' && Number.isInteger(token) && Number.isInteger(gen)) {
				const keyword = this.token();
				if (keyword instanceof PdfKeyword && keyword.word === 'R') return new PdfRef(token, gen);
			}
			this.pos = save;
			return token;
		}
		return token;
	}

	/** A dictionary value: an object, skipping stray keywords. */
	#value(depth: number): PdfValue | undefined {
		const value = this.#build(this.token(), depth);
		if (value === EOF) return undefined;
		if (value instanceof PdfKeyword) return null;
		return value;
	}
}

export function hexValue(byte: number): number {
	if (byte >= 0x30 && byte <= 0x39) return byte - 0x30;
	if (byte >= 0x41 && byte <= 0x46) return byte - 0x37;
	if (byte >= 0x61 && byte <= 0x66) return byte - 0x57;
	return -1;
}

export function latin1(bytes: Uint8Array, start = 0, end = bytes.length): string {
	let text = '';
	for (let at = start; at < end && at < bytes.length; at += 8192)
		text += String.fromCharCode(...bytes.subarray(at, Math.min(end, at + 8192)));
	return text;
}
