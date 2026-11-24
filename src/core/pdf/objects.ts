/** The PDF object model (ISO 32000-1 section 7.3): what the parser reads and the reader resolves. */

export class PdfName {
	constructor(readonly name: string) {}
}

/** A string object as its bytes (literal or hexadecimal; decrypted when the file is encrypted). */
export class PdfString {
	constructor(readonly bytes: Uint8Array) {}
}

export class PdfRef {
	constructor(
		readonly num: number,
		readonly gen: number,
	) {}
	get key(): string {
		return `${this.num} ${this.gen}`;
	}
}

export type PdfDict = Map<string, PdfValue>;

/** A stream: its dictionary and its still-encoded (and, in an encrypted file, encrypted) bytes. */
export class PdfStream {
	constructor(
		readonly dict: PdfDict,
		readonly raw: Uint8Array,
		/** The indirect object holding the stream, for its decryption key. */
		readonly ref?: PdfRef,
	) {}
}

export type PdfValue =
	| null
	| boolean
	| number
	| PdfName
	| PdfString
	| PdfRef
	| PdfStream
	| PdfDict
	| PdfValue[];

export const isDict = (value: unknown): value is PdfDict => value instanceof Map;
export const isName = (value: unknown, name?: string): value is PdfName =>
	value instanceof PdfName && (name === undefined || value.name === name);
export const isNumber = (value: unknown): value is number =>
	typeof value === 'number' && Number.isFinite(value);

/** A PDF string as text: UTF-16BE with a byte order mark, UTF-8 with one, else PDFDocEncoding. */
export function decodeTextString(value: PdfValue | undefined): string | undefined {
	if (!(value instanceof PdfString)) return undefined;
	const bytes = value.bytes;
	if (bytes[0] === 0xfe && bytes[1] === 0xff) {
		let text = '';
		for (let at = 2; at + 1 < bytes.length; at += 2)
			text += String.fromCharCode((bytes[at]! << 8) | bytes[at + 1]!);
		return text;
	}
	if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
		return new TextDecoder().decode(bytes.subarray(3));
	let text = '';
	for (const byte of bytes) text += PDF_DOC_HIGH[byte - 0x80] ?? String.fromCharCode(byte);
	return text;
}

/** PDFDocEncoding differs from Latin-1 in 0x80..0x9f (ISO 32000-1 annex D). */
const PDF_DOC_HIGH: (string | undefined)[] = [
	'\u2022',
	'\u2020',
	'\u2021',
	'\u2026',
	'\u2014',
	'\u2013',
	'\u0192',
	'\u2044',
	'\u2039',
	'\u203a',
	'\u2212',
	'\u2030',
	'\u201e',
	'\u201c',
	'\u201d',
	'\u2018',
	'\u2019',
	'\u201a',
	'\u2122',
	'\ufb01',
	'\ufb02',
	'\u0141',
	'\u0152',
	'\u0160',
	'\u0178',
	'\u017d',
	'\u0131',
	'\u0142',
	'\u0153',
	'\u0161',
	'\u017e',
	undefined,
];
