// Excel's case mapping and surrogate-aware slicing for the text functions.

const isLetter = (ch: string): boolean => ch.toLowerCase() !== ch.toUpperCase();

/** One code point upper-cased; mappings that expand ("ß" to "SS", "ﬁ" to "FI") keep the original. */
function upperChar(ch: string): string {
	const up = ch.toUpperCase();
	return up.length === ch.length || [...up].length === 1 ? up : ch;
}

/**
 * One code point lower-cased: an expanding mapping keeps its first code point ("İ" becomes "i"),
 * and capital sigma becomes final sigma at the end of a word, as Excel does.
 */
function lowerChar(ch: string, before: string | undefined, after: string | undefined): string {
	if (ch === 'Σ') {
		const final =
			before !== undefined && isLetter(before) && !(after !== undefined && isLetter(after));
		return final ? 'ς' : 'σ';
	}
	const low = ch.toLowerCase();
	const points = [...low];
	return points.length === 1 ? low : (points[0] ?? ch);
}

function mapChars(text: string, fn: (ch: string, i: number, all: string[]) => string): string {
	const chars = [...text];
	let out = '';
	for (let i = 0; i < chars.length; i++) out += fn(chars[i] as string, i, chars);
	return out;
}

/** UPPER: per code point, so the length only changes where Excel's does (never). */
export const upper = (text: string): string => mapChars(text, upperChar);

/** LOWER: per code point, with Excel's final-sigma handling. */
export const lower = (text: string): string =>
	mapChars(text, (ch, i, all) => lowerChar(ch, all[i - 1], all[i + 1]));

/** PROPER: a letter after a non-letter is upper-cased, every other letter lower-cased. */
export const proper = (text: string): string =>
	mapChars(text, (ch, i, all) => {
		if (!isLetter(ch)) return ch;
		const before = all[i - 1];
		return before !== undefined && isLetter(before)
			? lowerChar(ch, before, all[i + 1])
			: upperChar(ch);
	});

const isHigh = (code: number): boolean => code >= 0xd800 && code <= 0xdbff;
const isLow = (code: number): boolean => code >= 0xdc00 && code <= 0xdfff;

/**
 * LEFT counts UTF-16 code units like Excel's LEN, but never ends inside a surrogate pair: LEFT of
 * an emoji with 1 character keeps both halves.
 */
export function left(text: string, n: number): string {
	let end = Math.min(n, text.length);
	if (end > 0 && end < text.length && isHigh(text.charCodeAt(end - 1))) end++;
	return text.slice(0, end);
}

/** RIGHT, likewise never starting inside a surrogate pair. */
export function right(text: string, n: number): string {
	if (n <= 0) return '';
	let start = Math.max(0, text.length - n);
	if (start > 0 && isLow(text.charCodeAt(start))) start--;
	return text.slice(start);
}
