/**
 * Linear format build-up: UnicodeMath (Unicode Technical Note 28), the notation Word's equation
 * editor types in, read into the LaTeX subset `latex-to-omml` converts, so both input formats share
 * one converter. Covered: fractions (`a/b`, `a∕b`, `a¦b`), scripts (`x^2`, `x_i^2`, `x^2^3` nested
 * to the right), roots (`√x`, `√(n&x)`, `∛`, `∜`), n-ary operators with limits and `▒`, brackets
 * (including `├ ┤` and `│` separators), functions (`sin x`, `lim_(n→∞) x`), accents (combining
 * marks), `□ ¯ ▁ ⏞ ⏟`, limits above and below (`┴ ┬`), matrices (`■(a&b@c&d)`), equation arrays
 * (`█(a@b)`), normal text (`"..."`), placeholders (`⬚`), `\` before a character to keep it literal
 * and math AutoCorrect control words (`\alpha`). Text it cannot structure is kept as text; it
 * throws only when nesting is deeper than `MAX_LINEAR_NESTING`.
 */
import { ACCENT_ALIASES, REVERSE_ACCENT } from './index';
import {
	bracket,
	empty,
	escapeChar,
	grid,
	nary,
	normalText,
	root,
	wrapper,
	WRAPPERS,
	type LinearReader,
	type Operand,
} from './unicode-math-constructs';
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

/** Deepest nesting of groups and structures the build-up reads. */
export const MAX_LINEAR_NESTING = 128;

/** Replaces each known `\word` with its character; unknown words stay as typed. */
export function applyControlWords(text: string): string {
	return text.replace(/\\([A-Za-z]+)/g, (word, name: string) => MATH_AUTOCORRECT[name] ?? word);
}

const BUILD_UP = new Set([
	LINEAR.fraction,
	LINEAR.linearFraction,
	LINEAR.stack,
	LINEAR.below,
	LINEAR.above,
]);

class Parser implements LinearReader {
	private index = 0;
	private depth = 0;
	private readonly chars: string[];

	constructor(text: string) {
		this.chars = Array.from(applyControlWords(text));
	}

	peek(offset = 0): string | undefined {
		return this.chars[this.index + offset];
	}

	take(): string {
		return this.chars[this.index++] ?? '';
	}

	skipSpaces(): void {
		while (this.peek() === ' ' || this.peek() === LINEAR.funcApply) this.index++;
	}

	private nested<T>(read: () => T): T {
		if (++this.depth > MAX_LINEAR_NESTING)
			throw new Error(`Linear format nesting exceeds ${MAX_LINEAR_NESTING} levels`);
		try {
			return read();
		} finally {
			this.depth--;
		}
	}

	/** A sequence up to (not including) one of `stops` or an unmatched closing bracket. */
	sequence(stops: ReadonlySet<string>, closer?: string): string {
		return this.nested(() => {
			let out = '';
			while (this.index < this.chars.length) {
				const char = this.peek()!;
				if (stops.has(char) || (closer !== undefined && this.closes(char, closer))) break;
				if (char === ' ' || char === LINEAR.funcApply) {
					this.index++;
				} else if (this.startsOperand(char)) {
					out += this.fractionChain().latex;
				} else if (char === '_' || char === '^') {
					out += this.prescript();
				} else {
					this.index++;
					// `\` keeps the character after it literal (`km\/h`).
					const escaped = char === '\\' && /^[^\p{L}]$/u.test(this.peek() ?? '');
					const literal = escaped ? this.take() : char;
					const op = escapeChar(literal);
					// An operator takes a limit above or below it too (`=┴"def"`, `→┬n`).
					const next = this.peek();
					out +=
						next === LINEAR.above || next === LINEAR.below
							? this.fractionChain({ latex: op, bare: op }).latex
							: op;
				}
			}
			return out;
		});
	}

	closes(char: string, closer: string): boolean {
		if (char === '|' || char === '‖') return char === closer;
		return CLOSERS.has(char);
	}

	startsOperand(char: string): boolean {
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
		for (let op = this.peek(); op !== undefined && BUILD_UP.has(op as never); op = this.peek()) {
			this.index++;
			this.skipSpaces();
			const next = this.peek();
			const right = next !== undefined && this.startsOperand(next) ? this.scripted() : empty();
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
		return left;
	}

	private accents(base: Operand): Operand {
		while (this.peek() !== undefined && isCombining(this.peek()!)) {
			const mark = this.take();
			const chr = ACCENT_ALIASES[mark] ?? mark;
			const command = REVERSE_ACCENT[chr];
			const latex = command ? `${command}{${base.bare}}` : `\\overset{${chr}}{${base.bare}}`;
			base = { latex, bare: latex };
		}
		return base;
	}

	/** A primary with its accents and its subscript and superscript. */
	scripted(): Operand {
		return this.nested(() => {
			let base = this.accents(this.primary());
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
			const core = chars.slice(split).join('');
			const head = Array.from(core).length > 1 ? `{${core}}` : core;
			const latex = `${chars.slice(0, split).join('')}${head}${sub !== undefined ? `_{${sub}}` : ''}${sup !== undefined ? `^{${sup}}` : ''}`;
			return { latex, bare: latex };
		});
	}

	/** Scripts with no base before them belong to the operand after them (`_a^b x`). */
	private prescript(): string {
		const { sub, sup } = this.scripts();
		this.skipSpaces();
		const next = this.peek();
		const base = next !== undefined && this.startsOperand(next) ? this.scripted() : empty();
		return `{}_{${sub ?? ''}}^{${sup ?? ''}}{${base.bare}}`;
	}

	scripts(): { sub?: string; sup?: string } {
		const out: { sub?: string; sup?: string } = {};
		for (let char = this.peek(); char === '_' || char === '^'; char = this.peek()) {
			const key = char === '_' ? 'sub' : 'sup';
			if (out[key] !== undefined) return out;
			this.index++;
			out[key] = this.scriptArgument(char);
		}
		return out;
	}

	/**
	 * A script's argument: a sign may lead it (`e^-x`), its brackets are removed, and the same
	 * script after it nests inside it (`x^2^3` is x to the 2 to the 3, as in Word).
	 */
	private scriptArgument(op: string): string {
		return this.nested(() => {
			let sign = '';
			while (['-', '+', '−', '±'].includes(this.peek() ?? '')) sign += this.take();
			const next = this.peek();
			if (next === undefined || !this.startsOperand(next)) return sign;
			const base = this.accents(this.primary());
			let latex = base.bare;
			if (this.peek() === op) {
				this.index++;
				const inner = this.scriptArgument(op);
				const head = Array.from(latex).length > 1 ? `{${latex}}` : latex;
				latex = `${head}${op}{${inner}}`;
			}
			return sign + latex;
		});
	}

	private primary(): Operand {
		const char = this.peek()!;
		if (char === '"') return normalText(this);
		if (char === LINEAR.groupOpen) {
			this.index++;
			const inner = this.sequence(new Set([LINEAR.groupClose]));
			if (this.peek() === LINEAR.groupClose) this.index++;
			return { latex: `{${inner}}`, bare: inner };
		}
		if (char in BRACKETS) return bracket(this);
		if (char in ROOTS) return root(this);
		if (char in NARY_CHARS) return nary(this);
		if (char === LINEAR.matrix || char === LINEAR.eqArray) return grid(this);
		if (char in WRAPPERS) return wrapper(this);
		return this.word();
	}

	private word(): Operand {
		let run = '';
		while (this.peek() !== undefined && isOperandChar(this.peek()!)) {
			run += this.take();
			if (this.peek() === '.' && /\d$/.test(run) && /\d/.test(this.peek(1) ?? ''))
				run += this.take();
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
		const argument = next !== undefined && this.startsOperand(next) ? this.scripted().latex : '';
		const latex = `${head}{${argument}}`;
		return { latex, bare: latex };
	}
}

/**
 * LaTeX for linear (UnicodeMath) text, ready for `convertLatexToOmml`. Throws only when the text
 * nests deeper than `MAX_LINEAR_NESTING`.
 */
export function unicodeMathToLatex(text: string): string {
	return new Parser(text).sequence(new Set());
}
