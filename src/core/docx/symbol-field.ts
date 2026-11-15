// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// SYMBOL (ECMA-376 Part 1, 17.16.5.67): one character by its code, in a font and size of its own.
// The details follow Word's output (`__fixtures__/computed-fields/symbol-1.docx`): a decimal or
// `0x` hexadecimal code, a fraction cut off; without `\u` a code below 256 is a Windows-1252
// byte and a larger one must be a character that code page holds; `\u` takes the code's low 16
// bits; a symbol font draws codes below 256 from its U+F000 page; a size is a whole number of
// points (at most 1638). A code Word cannot draw shows `###`.
import { windows1252CodePoint, windows1252HighByte } from '../text/windows-1252';
import { parseFieldInstruction } from './field-format';

/** What a SYMBOL field shows. */
export interface SymbolFieldResult {
	text: string;
	/** `\f`: the font the character is drawn in. */
	fontFamily?: string;
	/** `\s`: its size in points. */
	fontSizePt?: number;
	/** `\h`: the character does not change its line's spacing (not applied by this layout). */
	keepLineSpacing?: true;
}

/** The result Word shows for a code it cannot draw. */
export const SYMBOL_ERROR = '###';

/** Fonts whose characters Word codes from U+F000 (symbol character set fonts). */
const SYMBOL_FONTS = new Set([
	'symbol',
	'wingdings',
	'wingdings 2',
	'wingdings 3',
	'webdings',
	'marlett',
	'mt extra',
]);

/** Whether `font` is a symbol font, whose codes below 256 are its U+F000 page. */
export const isSymbolFont = (font: string | undefined): boolean =>
	font !== undefined && SYMBOL_FONTS.has(font.trim().toLowerCase());

/** The code argument: decimal (a fraction cut off) or `0x` hexadecimal; NaN when it is neither. */
function code(argument: string): number {
	if (/^0x[0-9a-f]+$/i.test(argument)) return Number.parseInt(argument.slice(2), 16);
	return /^\d+(\.\d*)?$/.test(argument) ? Math.trunc(Number(argument)) : Number.NaN;
}

/** The character a code draws in `font`, or undefined when Word cannot draw it. */
function character(value: number, unicode: boolean, font: string | undefined): string | undefined {
	if (!Number.isFinite(value) || value <= 0) return undefined;
	const point = unicode ? value & 0xffff : value;
	if (point === 0) return undefined;
	if (isSymbolFont(font) && point < 0x100) return String.fromCharCode(0xf000 + point);
	if (unicode || isSymbolFont(font)) return String.fromCharCode(point);
	if (point < 0x100) return String.fromCharCode(windows1252CodePoint(point));
	return windows1252HighByte(point) === undefined ? undefined : String.fromCharCode(point);
}

/**
 * What a SYMBOL field shows, or undefined for another field or one with a switch this does not
 * read. `\a` (ANSI) and `\j` (Shift-JIS) read the code as Word showed them on a Western system:
 * as a Windows-1252 code.
 */
export function symbolFieldResult(instr: string): SymbolFieldResult | undefined {
	const field = parseFieldInstruction(instr, {
		f: 'arg',
		s: 'arg',
		u: 'flag',
		a: 'flag',
		j: 'flag',
		h: 'flag',
	});
	if (field?.name !== 'SYMBOL') return undefined;
	const font = field.switches.get('f');
	const fontFamily = typeof font === 'string' && font.trim() ? font : undefined;
	const text = character(code(field.args[0] ?? ''), field.switches.has('u'), fontFamily);
	if (text === undefined) return { text: SYMBOL_ERROR };
	const size = field.switches.get('s');
	const points =
		typeof size === 'string' && /^\d/.test(size) ? Math.trunc(Number.parseFloat(size)) : 0;
	return {
		text,
		...(fontFamily ? { fontFamily } : {}),
		...(points > 0 ? { fontSizePt: Math.min(points, 1638) } : {}),
		...(field.switches.has('h') ? { keepLineSpacing: true as const } : {}),
	};
}
