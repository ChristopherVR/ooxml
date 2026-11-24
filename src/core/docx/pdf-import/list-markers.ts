/**
 * List markers at the start of a line: bullets (typed characters and symbol-font glyphs) and
 * numbers ("1.", "1)", "(a)", "iv.", "1.2"), with where the item's text starts.
 */
import type { Line } from './lines';

export type ListFormat =
	| 'bullet'
	| 'decimal'
	| 'lowerLetter'
	| 'upperLetter'
	| 'lowerRoman'
	| 'upperRoman';

export interface ListMarker {
	format: ListFormat;
	/** The bullet glyph and the font it was drawn in. */
	bullet?: { char: string; font: string };
	prefix: string;
	suffix: string;
	/** The number (the last part of "1.2"); 1 for bullets. */
	value: number;
	/** Parts of a dotted number ("1.2" has 2); 1 otherwise. */
	depth: number;
	/** Characters of the line's text the marker and the space after it take. */
	length: number;
	/** A bullet glyph: a list item wherever it sits. Numbers and dashes need more evidence. */
	strong: boolean;
	/** Left edge of the marker and of the item's text. */
	x: number;
	textX: number;
	/** Space between the marker's ink and the text: a tab sets a marker off, a space does not. */
	gap: number;
}

type Parsed = Omit<ListMarker, 'length' | 'strong' | 'x' | 'textX' | 'gap'>;

const STRONG = new Set([
	...'\u2022\u25cf\u25e6\u25aa\u25ab\u25a0\u25a1\u00b7\u2219\u27a2\u25ba\u2023\u2043\u2713\u2714\u2756\u25c6\u25c7\u2666',
]);
const WEAK = new Set([...'-\u2013\u2014']);
const SYMBOL_FONT = /symbol|wingdings|webdings|dingbat/i;
const NUMBER =
	/^(\()?(\d{1,3}(?:\.\d{1,3}){0,3}|[a-zA-Z]|[ivxlcdm]{2,6}|[IVXLCDM]{2,6})([.)])?(?=\s)/;

function roman(text: string): number {
	const values: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
	let total = 0;
	const lower = text.toLowerCase();
	for (let at = 0; at < lower.length; at++) {
		const value = values[lower[at]!] ?? 0;
		const next = values[lower[at + 1] ?? ''] ?? 0;
		total += value < next ? -value : value;
	}
	return total;
}

function toRoman(value: number): string {
	const table: [number, string][] = [
		[1000, 'm'],
		[900, 'cm'],
		[500, 'd'],
		[400, 'cd'],
		[100, 'c'],
		[90, 'xc'],
		[50, 'l'],
		[40, 'xl'],
		[10, 'x'],
		[9, 'ix'],
		[5, 'v'],
		[4, 'iv'],
		[1, 'i'],
	];
	let out = '';
	for (const [size, letters] of table)
		while (value >= size) {
			out += letters;
			value -= size;
		}
	return out;
}

function parseNumber(text: string): Parsed | undefined {
	const match = NUMBER.exec(text);
	if (!match) return undefined;
	const [, open = '', token = '', close = ''] = match;
	if (open && close !== ')') return undefined;
	const dotted = /^\d+(\.\d+)+$/.test(token);
	if (!close && !dotted) return undefined;
	const base = { prefix: open, suffix: close, depth: 1 };
	if (/^\d/.test(token)) {
		const parts = token.split('.');
		return {
			...base,
			format: 'decimal',
			value: Number(parts[parts.length - 1]),
			depth: parts.length,
		};
	}
	if (token.length > 1 || token === 'i' || token === 'I') {
		const value = roman(token);
		if (!value || toRoman(value) !== token.toLowerCase()) return undefined;
		return { ...base, format: token === token.toLowerCase() ? 'lowerRoman' : 'upperRoman', value };
	}
	const lower = token === token.toLowerCase();
	return {
		...base,
		format: lower ? 'lowerLetter' : 'upperLetter',
		value: token.toLowerCase().charCodeAt(0) - 96,
	};
}

/** The list marker `line` starts with, if any. */
export function markerOf(line: Line): ListMarker | undefined {
	const first = line.fragments.findIndex((fragment) => fragment.text.trim());
	const fragment = line.fragments[first];
	if (!fragment) return undefined;
	const leading = line.text.length - line.text.trimStart().length;
	const text = line.text.slice(leading);
	const char = text[0] ?? '';
	const ownText = fragment.text.trim();
	const symbolFont = SYMBOL_FONT.test(fragment.font.family) && [...ownText].length === 1;
	const privateUse = /[\uf000-\uf0ff]/.test(char);
	let marker: Parsed | undefined;
	let token = '';
	let strong = false;
	if (
		STRONG.has(char) ||
		privateUse ||
		symbolFont ||
		WEAK.has(char) ||
		char === 'o' ||
		char === '*'
	) {
		token = symbolFont ? ownText : char;
		if (!/^\s/.test(text.slice(token.length)) && !(symbolFont || privateUse)) return undefined;
		strong = STRONG.has(char) || privateUse || symbolFont;
		marker = {
			format: 'bullet',
			bullet: { char: token, font: fragment.font.family },
			prefix: '',
			suffix: '',
			value: 1,
			depth: 1,
		};
	} else {
		marker = parseNumber(text);
		if (!marker) return undefined;
		token = `${marker.prefix}${NUMBER.exec(text)![2]}${marker.suffix}`;
	}
	const spaces = /^\s*/.exec(text.slice(token.length))![0].length;
	const length = leading + token.length + spaces;
	// Where the text starts: the next fragment when the marker is drawn on its own.
	let textX: number;
	let inkEnd: number;
	const offset = fragment.text.length - fragment.text.trimStart().length;
	const at = (chars: number) => fragment.x + (fragment.width * chars) / fragment.text.length;
	if (ownText === token) {
		const next = line.fragments.slice(first + 1).find((item) => item.text.trim());
		inkEnd = at(offset + token.length);
		textX = next?.x ?? fragment.x + fragment.width;
	} else {
		inkEnd = at(offset + token.length);
		textX = at(offset + token.length + spaces);
	}
	const gap = textX - inkEnd;
	// A lone "o" or "*" is a bullet only when set off like one (Word's level-2 "o").
	if ((char === 'o' || char === '*') && !symbolFont && gap < line.size * 0.6) return undefined;
	// So is a dotted number without a full stop ("1.2"): otherwise "3.5 million" would start a list.
	if (marker.depth > 1 && !marker.suffix && gap < line.size * 0.6) return undefined;
	return { ...marker, length, strong, x: fragment.x, textX, gap };
}
