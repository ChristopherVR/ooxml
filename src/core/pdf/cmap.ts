/**
 * CMaps (ISO 32000-1 section 9.7.5 and 9.10.3): the code space that splits a string into
 * character codes, codes to CIDs for composite fonts, and codes to Unicode (`/ToUnicode`).
 */
import { EOF, PdfKeyword, PdfLexer } from './lexer';
import { PdfName, PdfString } from './objects';

interface Range {
	low: number;
	high: number;
	bytes: number;
}

const MAX_RANGE = 0x10000;

const codeOf = (bytes: Uint8Array) => bytes.reduce((value, byte) => value * 256 + byte, 0);

function utf16(bytes: Uint8Array): string {
	let text = '';
	for (let at = 0; at + 1 < bytes.length; at += 2)
		text += String.fromCharCode((bytes[at]! << 8) | bytes[at + 1]!);
	if (bytes.length % 2) text += String.fromCharCode(bytes[bytes.length - 1]!);
	return text;
}

export class PdfCMap {
	readonly codespace: Range[] = [];
	readonly unicode = new Map<number, string>();
	readonly cid = new Map<number, number>();
	/** Code width when no code space is declared. */
	fallbackBytes = 2;
	/** Identity: the code is the CID. */
	identity = false;

	static identityH(): PdfCMap {
		const cmap = new PdfCMap();
		cmap.identity = true;
		cmap.codespace.push({ low: 0, high: 0xffff, bytes: 2 });
		return cmap;
	}

	/** Splits a string into character codes with their byte widths. */
	codes(bytes: Uint8Array): { code: number; bytes: number }[] {
		const out: { code: number; bytes: number }[] = [];
		for (let at = 0; at < bytes.length;) {
			let matched: number | undefined;
			if (this.codespace.length) {
				for (
					let width = 1;
					width <= 4 && matched === undefined && at + width <= bytes.length;
					width++
				) {
					const code = codeOf(bytes.subarray(at, at + width));
					if (
						this.codespace.some(
							(range) => range.bytes === width && code >= range.low && code <= range.high,
						)
					)
						matched = width;
				}
			}
			const width =
				matched ?? Math.min(this.codespace.length ? 1 : this.fallbackBytes, bytes.length - at);
			out.push({ code: codeOf(bytes.subarray(at, at + width)), bytes: width });
			at += width;
		}
		return out;
	}

	cidOf(code: number): number {
		if (this.identity) return code;
		return this.cid.get(code) ?? code;
	}

	/** Parses a CMap or ToUnicode stream (the PostScript subset CMaps use). */
	static parse(data: Uint8Array): PdfCMap {
		const cmap = new PdfCMap();
		const lexer = new PdfLexer(data);
		const operands: unknown[] = [];
		let guard = 0;
		for (;;) {
			if (++guard > 2_000_000) break;
			const value = lexer.object();
			if (value === EOF) break;
			if (!(value instanceof PdfKeyword)) {
				operands.push(value);
				continue;
			}
			const word = value.word;
			if (word === 'endcodespacerange') {
				for (let at = 0; at + 1 < operands.length; at += 2) {
					const low = operands[at];
					const high = operands[at + 1];
					if (low instanceof PdfString && high instanceof PdfString)
						cmap.codespace.push({
							low: codeOf(low.bytes),
							high: codeOf(high.bytes),
							bytes: low.bytes.length,
						});
				}
			} else if (word === 'endbfchar') {
				for (let at = 0; at + 1 < operands.length; at += 2) {
					const source = operands[at];
					const target = operands[at + 1];
					if (!(source instanceof PdfString)) continue;
					const text =
						target instanceof PdfString
							? utf16(target.bytes)
							: target instanceof PdfName
								? target.name
								: undefined;
					if (text !== undefined) cmap.unicode.set(codeOf(source.bytes), text);
				}
			} else if (word === 'endbfrange') {
				for (let at = 0; at + 2 < operands.length; at += 3) {
					const [low, high, target] = [operands[at], operands[at + 1], operands[at + 2]];
					if (!(low instanceof PdfString) || !(high instanceof PdfString)) continue;
					const from = codeOf(low.bytes);
					const to = Math.min(codeOf(high.bytes), from + MAX_RANGE);
					for (let code = from; code <= to; code++) {
						if (Array.isArray(target)) {
							const item = target[code - from];
							if (item instanceof PdfString) cmap.unicode.set(code, utf16(item.bytes));
						} else if (target instanceof PdfString) {
							// The last byte increments over the range.
							const bytes = target.bytes.slice();
							bytes[bytes.length - 1] = bytes[bytes.length - 1]! + (code - from);
							cmap.unicode.set(code, utf16(bytes));
						}
					}
				}
			} else if (word === 'endcidchar') {
				for (let at = 0; at + 1 < operands.length; at += 2) {
					const source = operands[at];
					const cid = operands[at + 1];
					if (source instanceof PdfString && typeof cid === 'number')
						cmap.cid.set(codeOf(source.bytes), cid);
				}
			} else if (word === 'endcidrange') {
				for (let at = 0; at + 2 < operands.length; at += 3) {
					const [low, high, cid] = [operands[at], operands[at + 1], operands[at + 2]];
					if (
						!(low instanceof PdfString) ||
						!(high instanceof PdfString) ||
						typeof cid !== 'number'
					)
						continue;
					const from = codeOf(low.bytes);
					const to = Math.min(codeOf(high.bytes), from + MAX_RANGE);
					for (let code = from; code <= to; code++) cmap.cid.set(code, cid + code - from);
				}
			}
			operands.length = 0;
		}
		return cmap;
	}
}
