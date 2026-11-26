// Small rules on the equation editor's linear text: where normal text is around the caret, and how
// a selection goes into the first placeholder of a structure inserted around it.
import { EQUATION_PLACEHOLDER, type EquationInputFormat } from 'ooxml-core/math';

/** Where `text` stands at `caret` inside normal text (`"..."` or `\text{...}`), if it does. */
export function normalTextAround(
	text: string,
	caret: number,
	format: EquationInputFormat,
): [number, number] | undefined {
	if (format === 'unicode') {
		const before = text.slice(0, caret);
		const quotes = (before.match(/(?<!\\)"/g) ?? []).length;
		if (quotes % 2 === 0) return undefined;
		const start = before.lastIndexOf('"');
		const end = text.indexOf('"', caret);
		return end < 0 ? undefined : [start, end + 1];
	}
	const start = text.lastIndexOf('\\text{', caret);
	if (start < 0) return undefined;
	const end = text.indexOf('}', start);
	return end >= caret - 1 && end >= start + 6 && caret > start ? [start, end + 1] : undefined;
}

const OPENS = new Set([...'([{|‖⟨⌊⌈〖&@│']);
const CLOSES = new Set([...')]}|‖⟩⌋⌉〗&@│']);

/**
 * The selected text as it goes into the first placeholder of `template`. In UnicodeMath more than
 * one operand is bracketed to stay one argument, unless the placeholder already sits inside a
 * bracket or a matrix cell (a Bracket structure must not double its brackets).
 */
export function placeholderArgument(
	template: string,
	selected: string,
	format: EquationInputFormat,
): string {
	if (format === 'latex' || /^[\p{L}\p{N}]+$/u.test(selected)) return selected;
	const chars = Array.from(template);
	const at = chars.indexOf(EQUATION_PLACEHOLDER);
	const enclosed = OPENS.has(chars[at - 1] ?? '') && CLOSES.has(chars[at + 1] ?? '');
	return enclosed ? selected : `(${selected})`;
}

/**
 * A structure's linear text with each placeholder that is the argument of a build-up operator
 * grouped (`〖⬚〗/〖⬚〗`), so whatever is typed in its place stays that argument (`a+b` in a
 * numerator) when the text builds up again. Placeholders already in a bracket or a cell stay bare.
 */
export function groupPlaceholders(text: string): string {
	const chars = Array.from(text);
	return chars
		.map((char, at) =>
			char === EQUATION_PLACEHOLDER &&
			!(OPENS.has(chars[at - 1] ?? '') && CLOSES.has(chars[at + 1] ?? ''))
				? `〖${char}〗`
				: char,
		)
		.join('');
}
