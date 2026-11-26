/**
 * The bracketed and prefixed constructs of the linear format (`unicode-math.ts`): brackets with
 * `│` separators, roots, n-ary operators, matrices and equation arrays, normal text and the
 * prefix wrappers (`□ ¯ ▁ ⏞ ⏟`). Each reads from the parser it is given and returns LaTeX.
 */
import { REVERSE_ESCAPE } from './index';
import { BRACKETS, LINEAR, NARY_CHARS, ROOTS } from './unicode-math-symbols';

/** An operand: its LaTeX, and its content without the brackets a build-up operator removes. */
export interface Operand {
	latex: string;
	bare: string;
	/** A run of letters and digits, whose script binds to its last letter. */
	run?: true;
}

/** What the constructs read their input through. */
export interface LinearReader {
	peek(offset?: number): string | undefined;
	take(): string;
	skipSpaces(): void;
	sequence(stops: ReadonlySet<string>, closer?: string): string;
	scripted(): Operand;
	scripts(): { sub?: string; sup?: string };
	startsOperand(char: string): boolean;
	closes(char: string, closer: string): boolean;
}

const DELIMITER_LATEX: Record<string, string> = {
	'{': '\\{',
	'}': '\\}',
	'⟨': '\\langle ',
	'⟩': '\\rangle ',
	'⌊': '\\lfloor ',
	'⌋': '\\rfloor ',
	'⌈': '\\lceil ',
	'⌉': '\\rceil ',
	'‖': '\\|',
	[LINEAR.invisibleOpen]: '.',
	[LINEAR.invisibleClose]: '.',
};

/** Prefix operators that wrap the operand after them. */
export const WRAPPERS: Record<string, string> = {
	[LINEAR.box]: '\\boxed',
	[LINEAR.overbar]: '\\overline',
	[LINEAR.underbar]: '\\underline',
	[LINEAR.overbrace]: '\\overbrace',
	[LINEAR.underbrace]: '\\underbrace',
};

export const empty = (): Operand => ({ latex: '', bare: '' });
const built = (latex: string): Operand => ({ latex, bare: latex });

/** A character as LaTeX text. */
export function escapeChar(char: string): string {
	if (char === '~') return '∼';
	if (char === '\\') return '\\backslash ';
	if (char === '&') return '\\&';
	return REVERSE_ESCAPE[char] ?? char;
}

/** The argument of a prefix operator: the next operand without its brackets. */
export function argument(reader: LinearReader): string {
	reader.skipSpaces();
	const next = reader.peek();
	return next === undefined || !reader.startsOperand(next) ? '' : reader.scripted().bare;
}

export function wrapper(reader: LinearReader): Operand {
	const command = WRAPPERS[reader.take()]!;
	return built(`${command}{${argument(reader)}}`);
}

/** `"..."`: normal text, where `\"` is a quote. */
export function normalText(reader: LinearReader): Operand {
	reader.take();
	let text = '';
	for (let char = reader.peek(); char !== undefined && char !== '"'; char = reader.peek()) {
		reader.take();
		if (char === '\\' && reader.peek() === '"') text += reader.take();
		else text += char;
	}
	if (reader.peek() === '"') reader.take();
	return built(
		`\\text{${text.replace(/[\\{}]/g, (c) => (c === '\\' ? '\\backslash ' : `\\${c}`))}}`,
	);
}

/** A bracketed group; `│` separates its elements. Parentheses are removed by build-up operators. */
export function bracket(reader: LinearReader): Operand {
	const open = reader.take();
	const expected = BRACKETS[open]!;
	const parts = [reader.sequence(new Set(['│']), expected)];
	while (reader.peek() === '│') {
		reader.take();
		parts.push(reader.sequence(new Set(['│']), expected));
	}
	const close = reader.peek();
	let closing: string = LINEAR.invisibleClose;
	if (close !== undefined && reader.closes(close, expected)) closing = reader.take();
	const inner = parts.join('\\middle|');
	const latex = `\\left${DELIMITER_LATEX[open] ?? open}${inner}\\right${DELIMITER_LATEX[closing] ?? closing}`;
	return { latex, bare: open === '(' && closing === ')' && parts.length === 1 ? inner : latex };
}

/** `√x`, `√(n&x)`, `∛x`, `∜x`. */
export function root(reader: LinearReader): Operand {
	let degree = ROOTS[reader.take()]!;
	reader.skipSpaces();
	let radicand: string;
	if (reader.peek() === '(') {
		reader.take();
		const first = reader.sequence(new Set(['&']), ')');
		if (reader.peek() === '&') {
			reader.take();
			degree = first;
			radicand = reader.sequence(new Set(), ')');
		} else radicand = first;
		if (reader.peek() === ')') reader.take();
	} else radicand = argument(reader);
	return built(degree ? `\\sqrt[${degree}]{${radicand}}` : `\\sqrt{${radicand}}`);
}

/** An n-ary operator, its limits and, after an optional `▒`, its operand. */
export function nary(reader: LinearReader): Operand {
	const command = NARY_CHARS[reader.take()]!;
	const { sub, sup } = reader.scripts();
	reader.skipSpaces();
	if (reader.peek() === LINEAR.naryand) reader.take();
	reader.skipSpaces();
	const next = reader.peek();
	let body = '';
	if (next !== undefined && reader.startsOperand(next)) {
		const operand = reader.scripted();
		body = next === LINEAR.groupOpen ? operand.bare : operand.latex;
	}
	const limits = `${sub !== undefined ? `_{${sub}}` : ''}${sup !== undefined ? `^{${sup}}` : ''}`;
	return built(`${command}${limits}{${body}}`);
}

/** `■(a&b@c&d)` or `█(a@b)`: rows split by `@`, cells by `&`. */
export function grid(reader: LinearReader): Operand {
	const kind = reader.take();
	reader.skipSpaces();
	const rows: string[][] = [];
	if (reader.peek() === '(') {
		reader.take();
		let row: string[] = [];
		for (;;) {
			row.push(reader.sequence(new Set(['&', '@']), ')'));
			const sep = reader.peek();
			if (sep === '&') reader.take();
			else if (sep === '@') {
				reader.take();
				rows.push(row);
				row = [];
			} else break;
		}
		rows.push(row);
		if (reader.peek() === ')') reader.take();
	}
	const env = kind === LINEAR.matrix ? 'matrix' : 'aligned';
	return built(`\\begin{${env}}${rows.map((cells) => cells.join('&')).join('\\\\')}\\end{${env}}`);
}
