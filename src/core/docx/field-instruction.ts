// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The words of a field instruction (`w:instrText`, `w:fldSimple/@w:instr`) and the general text
// formatting switches, per ECMA-376 Part 1, 17.16.1 (syntax) and 17.16.4.1 (format switches).

/** One word of a field instruction: text (an argument, unquoted) or a switch (`l` for `\l`). */
export interface FieldToken {
	kind: 'text' | 'switch';
	value: string;
}

/**
 * Splits an instruction into its field type, arguments and switches. A quoted argument keeps its
 * spaces, and `\"` or `\\` inside quotes stand for the character itself. A switch is a backslash
 * followed by `*`, `#`, `@`, `!` or one or two letters; its letters are lower-cased.
 */
export function fieldTokens(instr: string): FieldToken[] {
	const tokens: FieldToken[] = [];
	let index = 0;
	while (index < instr.length) {
		const char = instr[index]!;
		if (/\s/.test(char)) {
			index++;
			continue;
		}
		if (char === '"') {
			let value = '';
			index++;
			while (index < instr.length && instr[index] !== '"') {
				if (instr[index] === '\\' && /["\\]/.test(instr[index + 1] ?? '')) index++;
				value += instr[index++];
			}
			index++;
			tokens.push({ kind: 'text', value });
			continue;
		}
		const switchMatch = /^\\([*#@!]|[A-Za-z]{1,2}(?![A-Za-z]))/.exec(instr.slice(index));
		if (switchMatch) {
			tokens.push({ kind: 'switch', value: switchMatch[1]!.toLowerCase() });
			index += switchMatch[0].length;
			continue;
		}
		const word = /^\S+/.exec(instr.slice(index))![0];
		tokens.push({ kind: 'text', value: word });
		index += word.length;
	}
	return tokens;
}

/** The text case a `\*` switch asks for. */
export type FieldTextCase = 'upper' | 'lower' | 'firstcap' | 'caps';

const TEXT_CASES = new Set<string>(['upper', 'lower', 'firstcap', 'caps']);

/** Whether a `\*` switch's argument (any case) is one of the text case formats. */
export const isFieldTextCase = (format: string): format is FieldTextCase =>
	TEXT_CASES.has(format.toLowerCase());

/**
 * Applies a text case format to a field result: `Upper` and `Lower` change every letter,
 * `FirstCap` capitalizes the first letter of the first word and `Caps` that of every word. The
 * other letters are left as they are, as the standard words it.
 */
export function applyFieldTextCase(text: string, format: FieldTextCase): string {
	if (format === 'upper') return text.toUpperCase();
	if (format === 'lower') return text.toLowerCase();
	const capitalize = (word: string) =>
		word.replace(
			/^(\P{L}*)(\p{L})/u,
			(_match, lead: string, letter: string) => `${lead}${letter.toUpperCase()}`,
		);
	if (format === 'firstcap')
		return text.replace(
			/^(\s*)(\S+)/,
			(_match, space: string, word: string) => `${space}${capitalize(word)}`,
		);
	return text.replace(/\S+/g, capitalize);
}
