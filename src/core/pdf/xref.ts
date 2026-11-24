/**
 * The cross-reference data of a PDF (ISO 32000-1 sections 7.5.4 to 7.5.8): classic tables,
 * cross-reference streams and hybrid files, following `/Prev` and `/XRefStm` with the newest entry
 * winning. A damaged or missing table falls back to a scan for `n g obj` headers, as viewers do.
 */
import { decodeFilters } from './filters';
import { EOF, latin1, PdfKeyword, PdfLexer } from './lexer';
import { isDict, isNumber, PdfStream, type PdfDict, type PdfValue } from './objects';

export type XrefEntry =
	| { type: 'offset'; offset: number; gen: number }
	| { type: 'compressed'; stream: number; index: number };

export interface XrefTable {
	entries: Map<number, XrefEntry>;
	trailer: PdfDict;
}

const STREAM_LIMIT = 64 * 1024 * 1024;

/** Reads `n g obj <value> [stream ... endstream]` at `offset`; the stream's length comes from `length`. */
export function readIndirect(
	bytes: Uint8Array,
	offset: number,
	length: (value: PdfValue | undefined) => number | undefined,
): { num: number; gen: number; value: PdfValue } | undefined {
	const lexer = new PdfLexer(bytes, offset);
	const num = lexer.token();
	const gen = lexer.token();
	const keyword = lexer.token();
	if (typeof num !== 'number' || typeof gen !== 'number') return undefined;
	if (!(keyword instanceof PdfKeyword) || keyword.word !== 'obj') return undefined;
	const value = lexer.object();
	if (value === EOF || value instanceof PdfKeyword) return { num, gen, value: null };
	if (isDict(value)) {
		const save = lexer.pos;
		const next = lexer.token();
		if (next instanceof PdfKeyword && next.word === 'stream') {
			let start = lexer.pos;
			if (bytes[start] === 13) start++;
			if (bytes[start] === 10) start++;
			let size = length(value.get('Length'));
			const declaredEnd = size === undefined ? -1 : start + size;
			// Trust /Length only when `endstream` follows it; else search for the keyword.
			if (
				size === undefined ||
				!latin1(bytes, declaredEnd, declaredEnd + 20).includes('endstream')
			) {
				const end = indexOf(bytes, 'endstream', start);
				size = Math.max(0, (end < 0 ? bytes.length : end) - start);
				while (size > 0 && (bytes[start + size - 1] === 10 || bytes[start + size - 1] === 13))
					size--;
			}
			return { num, gen, value: new PdfStream(value, bytes.subarray(start, start + size)) };
		}
		lexer.pos = save;
	}
	return { num, gen, value };
}

export function indexOf(bytes: Uint8Array, text: string, from = 0): number {
	const first = text.charCodeAt(0);
	outer: for (let at = from; at <= bytes.length - text.length; at++) {
		if (bytes[at] !== first) continue;
		for (let index = 1; index < text.length; index++)
			if (bytes[at + index] !== text.charCodeAt(index)) continue outer;
		return at;
	}
	return -1;
}

function lastIndexOf(bytes: Uint8Array, text: string): number {
	for (let at = bytes.length - text.length; at >= 0; at--) {
		let match = true;
		for (let index = 0; index < text.length && match; index++)
			match = bytes[at + index] === text.charCodeAt(index);
		if (match) return at;
	}
	return -1;
}

function readTable(bytes: Uint8Array, offset: number, table: XrefTable): PdfDict | undefined {
	const lexer = new PdfLexer(bytes, offset);
	const head = lexer.token();
	if (!(head instanceof PdfKeyword) || head.word !== 'xref') return undefined;
	for (;;) {
		const start = lexer.token();
		if (start instanceof PdfKeyword && start.word === 'trailer') break;
		const count = lexer.token();
		if (typeof start !== 'number' || typeof count !== 'number') return undefined;
		for (let index = 0; index < count; index++) {
			const entryOffset = lexer.token();
			const gen = lexer.token();
			const kind = lexer.token();
			if (typeof entryOffset !== 'number' || typeof gen !== 'number') return undefined;
			const num = start + index;
			if (kind instanceof PdfKeyword && kind.word === 'n' && !table.entries.has(num) && num > 0)
				table.entries.set(num, { type: 'offset', offset: entryOffset, gen });
			else if (!table.entries.has(num)) table.entries.set(num, { type: 'offset', offset: -1, gen });
		}
	}
	const trailer = lexer.object();
	return isDict(trailer) ? trailer : undefined;
}

function readStream(bytes: Uint8Array, offset: number, table: XrefTable): PdfDict | undefined {
	const indirect = readIndirect(bytes, offset, (value) => (isNumber(value) ? value : undefined));
	const stream = indirect?.value;
	if (!(stream instanceof PdfStream)) return undefined;
	const { dict } = stream;
	const widths = dict.get('W');
	if (!Array.isArray(widths) || widths.length < 3 || !widths.every(isNumber)) return undefined;
	const [w1, w2, w3] = widths as number[];
	const size = isNumber(dict.get('Size')) ? (dict.get('Size') as number) : 0;
	const index = Array.isArray(dict.get('Index')) ? (dict.get('Index') as number[]) : [0, size];
	const data = decodeFilters(stream.raw, dict, STREAM_LIMIT).bytes;
	const field = (at: number, width: number) => {
		let value = 0;
		for (let byte = 0; byte < width; byte++) value = value * 256 + (data[at + byte] ?? 0);
		return value;
	};
	let at = 0;
	const rowWidth = w1! + w2! + w3!;
	for (let pair = 0; pair + 1 < index.length; pair += 2) {
		const first = index[pair]!;
		const count = index[pair + 1]!;
		for (let item = 0; item < count && at + rowWidth <= data.length; item++, at += rowWidth) {
			const type = w1 ? field(at, w1) : 1;
			const a = field(at + w1!, w2!);
			const b = field(at + w1! + w2!, w3!);
			const num = first + item;
			if (table.entries.has(num) || num === 0) continue;
			if (type === 1) table.entries.set(num, { type: 'offset', offset: a, gen: b });
			else if (type === 2) table.entries.set(num, { type: 'compressed', stream: a, index: b });
			else table.entries.set(num, { type: 'offset', offset: -1, gen: b });
		}
	}
	return dict;
}

/** Reads the cross-reference chain from `startxref`; `undefined` when it is damaged. */
function readChain(bytes: Uint8Array): XrefTable | undefined {
	const at = lastIndexOf(bytes, 'startxref');
	if (at < 0) return undefined;
	const lexer = new PdfLexer(bytes, at + 9);
	let offset = lexer.token();
	const table: XrefTable = { entries: new Map(), trailer: new Map() };
	const seen = new Set<number>();
	let first = true;
	while (typeof offset === 'number' && offset >= 0 && offset < bytes.length && !seen.has(offset)) {
		seen.add(offset);
		const trailer = readTable(bytes, offset, table) ?? readStream(bytes, offset, table);
		if (!trailer) return first ? undefined : table;
		if (first) table.trailer = trailer;
		first = false;
		const hybrid = trailer.get('XRefStm');
		if (isNumber(hybrid)) readStream(bytes, hybrid, table);
		const prev = trailer.get('Prev');
		offset = isNumber(prev) ? prev : -1;
	}
	return table.trailer.has('Root') ? table : undefined;
}

/** Rebuilds the table by scanning the file for object headers (a damaged file). */
function scan(bytes: Uint8Array): XrefTable {
	const text = latin1(bytes);
	const table: XrefTable = { entries: new Map(), trailer: new Map() };
	for (const match of text.matchAll(/(?:^|[\s>])(\d+)\s+(\d+)\s+obj\b/g)) {
		const offset = match.index + match[0].indexOf(match[1]!);
		table.entries.set(Number(match[1]), { type: 'offset', offset, gen: Number(match[2]) });
	}
	for (const match of text.matchAll(/trailer\s*<</g)) {
		const trailer = new PdfLexer(bytes, match.index + 7).object();
		if (isDict(trailer)) for (const [key, value] of trailer) table.trailer.set(key, value);
	}
	return table;
}

export function readXref(bytes: Uint8Array): XrefTable {
	try {
		const table = readChain(bytes);
		if (table) return table;
	} catch {
		// fall through to the scan
	}
	return scan(bytes);
}
