/**
 * Linear format build-up: UnicodeMath (Unicode Technical Note 28), the notation Word's equation
 * editor types in, read into the LaTeX subset `latex-to-omml` converts, so both input formats share
 * one converter. Covered: fractions (`a/b`, `a¦b`), scripts (`x^2`, `x_i`, `x_i^2`), roots (`√x`,
 * `√(n&x)`, `∛`, `∜`), n-ary operators with limits and `▒`, brackets (including `├ ┤`),
 * functions (`sin x`, `lim_(n→∞) x`), accents (combining marks), `□ ¯ ▁ ⏞ ⏟`, limits above and
 * below (`┴ ┬`), matrices (`■(a&b@c&d)`), equation arrays (`█(a@b)`), normal text (`"..."`),
 * placeholders (`⬚`) and math AutoCorrect control words (`\alpha`). Never throws: text it cannot
 * structure is kept as text.
 */
import { ACCENT_ALIASES, REVERSE_ACCENT, REVERSE_ESCAPE } from './latex-omml-symbols';
import {
	BRACKETS,
	CLOSERS,
	isCombining,
	isOperandChar,
	LINEAR,
	LINEAR_FUNCTIONS,
	MATH_AUTOCORRECT,
	NARY_CHARS,
	ROOTS,
} from './unicode-math-symbols';

const DELIMITER_LATEX: Record<string, string> = {
	'(': '(',
	')': ')',
	'[': '[',
	']': ']',
	'{': '\\{',
	'}': '\\}',
	'⟨': '\\langle ',
	'⟩': '\\rangle ',
	'⌊': '\\lfloor ',
	'⌋': '\\rfloor ',
	'⌈': '\\lceil ',
	'⌉': '\\rceil ',
	'|': '|',
	'‖': '\\|',
	[LINEAR.invisibleOpen]: '.',
	[LINEAR.invisibleClose]: '.',
};

/** Prefix operators that wrap the operand after them. */
const WRAPPERS: Record<string, string> = {
	[LINEAR.box]: '\\boxed',
	[LINEAR.overbar]: '\\overline',
	[LINEAR.underbar]: '\\underline',
	[LINEAR.overbrace]: '\\overbrace',
	[LINEAR.underbrace]: '\\underbrace',
};

/** An operand: its LaTeX, and its content without the brackets a build-up operator removes. */
interface Operand {
	latex: string;
	bare: string;
	/** A run of letters and digits, whose script binds to its last letter. */
	run?: true;
}

/** Replaces each known `\word` with its character; unknown words stay as typed. */
export function applyControlWords(text: string): string {
	return text.replace(/\\([A-Za-z]+)/g, (word, name: string) => MATH_AUTOCORRECT[name] ?? word);
}

function escapeChar(char: string): string {
	if (char === '~') return '∼';
	if (char === '\\') return '\\backslash ';
	if (char === '&') return '\\&';
	return REVERSE_ESCAPE[char] ?? char;
}

class Parser {
	private index = 0;
	private readonly chars: string[];

	constructor(text: string) {
		this.chars = Array.from(applyControlWords(text));
	}

	private peek(offset = 0): string | undefined {
		return this.chars[this.index + offset];
	}

	private skipSpaces(): void {
		while (this.peek() === ' ' || this.peek() === LINEAR.funcApply) this.index++;
	}

	/** A sequence up to (not including) one of `stops` or an unmatched closing bracket. */
	sequence(stops: ReadonlySet<string>, closer?: string): string {
		let out = '';
		while (this.index < this.chars.length) {
			const char = this.peek()!;
			if (stops.has(char) || (closer !== undefined && this.closes(char, closer))) break;
			if (char === ' ' || char === LINEAR.funcApply) {
				this.index++;
				continue;
			}
			if (this.startsOperand(char)) {
				out += this.fractionChain().latex;
				continue;
			}
			this.index++;
			const next = this.peek();
			if (char === '_' || char === '^') {
				out += this.prescript();
				continue;
			}
			// An operator takes a limit above or below it too (`=┴"def"`, `→┬n`).
			if (next === LINEAR.above || next === LINEAR.below) {
				const op = escapeChar(char);
				out += this.fractionChain({ latex: op, bare: op }).latex;
				continue;
			}
			out += escapeChar(char);
		}
		return out;
	}

	private closes(char: string, closer: string): boolean {
		if (char === '|' || char === '‖') return char === closer;
		return CLOSERS.has(char);
	}

	private startsOperand(char: string): boolean {
		return (
			isOperandChar(char) ||
			char in BRACKETS ||
			char === LINEAR.groupOpen ||
			char in ROOTS ||
			char in NARY_CHARS ||
			char === '"' ||
			char === LINEAR.matrix ||
			char === LINEAR.eqArray ||
			char in WRAPPERS
		);
	}

	/** An operand followed by fraction, stack and limit operators, which group to the left. */
	private fractionChain(first?: Operand): Operand {
		let left = first ?? this.scripted();
		for (;;) {
			const op = this.peek();
			if (
				op !== LINEAR.fraction &&
				op !== LINEAR.linearFraction &&
				op !== LINEAR.stack &&
				op !== LINEAR.below &&
				op !== LINEAR.above
			)
				return left;
			this.index++;
			this.skipSpaces();
			const right =
				this.peek() !== undefined && this.startsOperand(this.peek()!) ? this.scripted() : empty();
			const latex =
				op === LINEAR.fraction
					? `\\frac{${left.bare}}{${right.bare}}`
					: op === LINEAR.linearFraction
						? `\\sfrac{${left.bare}}{${right.bare}}`
						: op === LINEAR.stack
							? `{${left.bare} \\atop ${right.bare}}`
							: `\\${op === LINEAR.below ? 'underset' : 'overset'}{${right.bare}}{${left.bare}}`;
			left = { latex, bare: latex };
		}
	}

	/** A primary with its accents and its subscript and superscript. */
	private scripted(): Operand {
		let base = this.primary();
		while (this.peek() !== undefined && isCombining(this.peek()!)) {
			const mark = this.chars[this.index++]!;
			const chr = ACCENT_ALIASES[mark] ?? mark;
			const command = REVERSE_ACCENT[chr];
			const latex = command ? `${command}{${base.bare}}` : `\\overset{${chr}}{${base.bare}}`;
			base = { latex, bare: latex };
		}
		if (this.peek() === "'") {
			let primes = '';
			while (this.peek() === "'") primes += (this.index++, '′');
			base = { latex: base.latex + primes, bare: base.latex + primes };
		}
		const { sub, sup } = this.scripts();
		if (sub === undefined && sup === undefined) return base;
		// `2x^2` raises only the x: a run's script binds to its last letter, as Word builds it up.
		const chars = Array.from(base.latex);
		const split = base.run && !/^[\d.]+$/.test(base.latex) ? chars.length - 1 : 0;
		const prefix = chars.slice(0, split).join('');
		const core = chars.slice(split).join('');
		const head = Array.from(core).length > 1 ? `{${core}}` : core;
		const latex = `${prefix}${head}${sub !== undefined ? `_{${sub}}` : ''}${sup !== undefined ? `^{${sup}}` : ''}`;
		return { latex, bare: latex };
	}

	/** Scripts with no base before them belong to the operand after them (`_a^b x`). */
	private prescript(): string {
		this.index--;
		const { sub, sup } = this.scripts();
		this.skipSpaces();
		const next = this.peek();
		const base = next !== undefined && this.startsOperand(next) ? this.scripted() : empty();
		return `{}_{${sub ?? ''}}^{${sup ?? ''}}{${base.bare}}`;
	}

	private scripts(): { sub?: string; sup?: string } {
		const out: { sub?: string; sup?: string } = {};
		for (;;) {
			const char = this.peek();
			if (char !== '_' && char !== '^') return out;
			const key = char === '_' ? 'sub' : 'sup';
			if (out[key] !== undefined) return out;
			this.index++;
			out[key] = this.scriptArgument();
		}
	}

	/** A script's argument: a sign may lead it (`e^-x`), its brackets are removed. */
	private scriptArgument(): string {
		let sign = '';
		while (this.peek() === '-' || this.peek() === '+' || this.peek() === '−' || this.peek() === '±')
			sign += this.chars[this.index++];
		const next = this.peek();
		if (next === undefined || !this.startsOperand(next)) return sign;
		const base = this.primary();
		let latex = base.bare;
		while (this.peek() !== undefined && isCombining(this.peek()!)) {
			const mark = this.chars[this.index++]!;
			const command = REVERSE_ACCENT[ACCENT_ALIASES[mark] ?? mark];
			latex = command ? `${command}{${latex}}` : latex;
		}
		return sign + latex;
	}

	private primary(): Operand {
		const char = this.peek()!;
		if (char === '"') return this.text();
		if (char === LINEAR.groupOpen) {
			this.index++;
			const inner = this.sequence(new Set([LINEAR.groupClose]));
			if (this.peek() === LINEAR.groupClose) this.index++;
			return { latex: `{${inner}}`, bare: inner };
		}
		if (char in BRACKETS) return this.bracket();
		if (char in ROOTS) return this.root();
		if (char in NARY_CHARS) return this.nary();
		if (char === LINEAR.matrix || char === LINEAR.eqArray) return this.grid(char);
		const wrapper = WRAPPERS[char];
		if (wrapper) {
			this.index++;
			const arg = this.argument();
			const latex = `${wrapper}{${arg}}`;
			return { latex, bare: latex };
		}
		return this.word();
	}

	/** The argument of a prefix operator: the next operand without its brackets. */
	private argument(): string {
		this.skipSpaces();
		const next = this.peek();
		if (next === undefined || !this.startsOperand(next)) return '';
		return this.scripted().bare;
	}

	private word(): Operand {
		let run = '';
		while (this.peek() !== undefined && isOperandChar(this.peek()!)) {
			run += this.chars[this.index++];
			if (this.peek() === '.' && /\d$/.test(run) && /\d/.test(this.peek(1) ?? ''))
				run += this.chars[this.index++];
		}
		if (LINEAR_FUNCTIONS.has(run)) return this.func(run);
		return { latex: run, bare: run, run: true };
	}

	/** A function name: its scripts, then the operand it applies to. */
	private func(name: string): Operand {
		const { sub, sup } = this.scripts();
		this.skipSpaces();
		const next = this.peek();
		const head = `\\${name}${sub !== undefined ? `_{${sub}}` : ''}${sup !== undefined ? `^{${sup}}` : ''}`;
		if (next === undefined || !this.startsOperand(next)) {
			const latex = `${head}{}`;
			return { latex, bare: latex };
		}
		const argument = this.scripted();
		const latex = `${head}{${argument.latex}}`;
		return { latex, bare: latex };
	}

	private text(): Operand {
		this.index++;
		let text = '';
		while (this.peek() !== undefined && this.peek() !== '"') text += this.chars[this.index++];
		if (this.peek() === '"') this.index++;
		const latex = `\\text{${text.replace(/[\\{}]/g, (c) => `\\${c === '\\' ? 'backslash ' : c}`)}}`;
		return { latex, bare: latex };
	}

	private bracket(): Operand {
		const open = this.chars[this.index++]!;
		const expected = BRACKETS[open]!;
		const inner = this.sequence(new Set(), expected);
		const close = this.peek();
		let closing = LINEAR.invisibleClose as string;
		if (close !== undefined && this.closes(close, expected)) {
			closing = close;
			this.index++;
		}
		const latex = `\\left${DELIMITER_LATEX[open] ?? open}${inner}\\right${DELIMITER_LATEX[closing] ?? closing}`;
		return { latex, bare: open === '(' && closing === ')' ? inner : latex };
	}

	private root(): Operand {
		const symbol = this.chars[this.index++]!;
		this.skipSpaces();
		let degree = ROOTS[symbol]!;
		let radicand = '';
		if (this.peek() === '(') {
			this.index++;
			const first = this.sequence(new Set(['&']), ')');
			if (this.peek() === '&') {
				this.index++;
				degree = first;
				radicand = this.sequence(new Set(), ')');
			} else radicand = first;
			if (this.peek() === ')') this.index++;
		} else radicand = this.argument();
		const latex = degree ? `\\sqrt[${degree}]{${radicand}}` : `\\sqrt{${radicand}}`;
		return { latex, bare: latex };
	}

	private nary(): Operand {
		const command = NARY_CHARS[this.chars[this.index++]!]!;
		const { sub, sup } = this.scripts();
		this.skipSpaces();
		if (this.peek() === LINEAR.naryand) this.index++;
		const body = this.argumentKeepingBrackets();
		const latex = `${command}${sub !== undefined ? `_{${sub}}` : ''}${sup !== undefined ? `^{${sup}}` : ''}{${body}}`;
		return { latex, bare: latex };
	}

	private argumentKeepingBrackets(): string {
		this.skipSpaces();
		const next = this.peek();
		if (next === undefined || !this.startsOperand(next)) return '';
		return next === LINEAR.groupOpen ? this.scripted().bare : this.scripted().latex;
	}

	/** `■(a&b@c&d)` or `█(a@b)`: rows split by `@`, cells by `&`. */
	private grid(kind: string): Operand {
		this.index++;
		this.skipSpaces();
		const rows: string[][] = [];
		if (this.peek() === '(') {
			this.index++;
			let row: string[] = [];
			for (;;) {
				row.push(this.sequence(new Set(['&', '@']), ')'));
				const sep = this.peek();
				if (sep === '&') this.index++;
				else if (sep === '@') {
					this.index++;
					rows.push(row);
					row = [];
				} else break;
			}
			rows.push(row);
			if (this.peek() === ')') this.index++;
		}
		const env = kind === LINEAR.matrix ? 'matrix' : 'aligned';
		const body = rows.map((cells) => cells.join('&')).join('\\\\');
		const latex = `\\begin{${env}}${body}\\end{${env}}`;
		return { latex, bare: latex };
	}
}

const empty = (): Operand => ({ latex: '', bare: '' });

/** LaTeX for linear (UnicodeMath) text, ready for `convertLatexToOmml`. */
export function unicodeMathToLatex(text: string): string {
	return new Parser(text).sequence(new Set());
}
