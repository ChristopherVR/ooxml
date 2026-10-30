// Word's wildcard syntax (Find > Use wildcards) translated to a regular-expression source. Every
// pattern Word accepts becomes an expression; every pattern Word rejects throws FindPatternError.
import { escapeFindChar, findCodeSource, FindPatternError } from './find-syntax';

const WORD_CHARACTERS = '\\p{L}\\p{M}\\p{N}_';
/** `?` and `*` match any character except an object or a paragraph mark. */
export const WILDCARD_ANY = '[^\\u0000\\r]';
const MAX_GROUPS = 9;

export interface WildcardSource {
	readonly source: string;
	readonly groupCount: number;
}

export function wildcardSource(query: string): WildcardSource {
	const parser = new WildcardParser(query);
	const source = parser.sequence(0);
	if (!parser.done) throw new FindPatternError('Unmatched ) in the search pattern');
	return { source, groupCount: parser.groupCount };
}

/** A character inside `[ ]`, escaped where the class syntax needs it. */
function classCharacter(char: string): string {
	return /[\\\]\[^-]/u.test(char) ? `\\${char}` : char;
}

class WildcardParser {
	private index = 0;
	groupCount = 0;

	constructor(private readonly query: string) {}

	get done(): boolean {
		return this.index >= this.query.length;
	}

	/** The next whole character (a surrogate pair is one character). */
	private take(): string {
		const code = this.query.codePointAt(this.index);
		if (code === undefined) throw new FindPatternError('The search pattern ends unexpectedly');
		const char = String.fromCodePoint(code);
		this.index += char.length;
		return char;
	}

	private peek(): string {
		return this.query.charAt(this.index);
	}

	/** The atoms up to a closing parenthesis, or the end of the pattern, as one expression. */
	sequence(depth: number): string {
		const atoms: string[] = [];
		while (!this.done) {
			if (this.peek() === ')') {
				if (depth === 0) throw new FindPatternError('Unmatched ) in the search pattern');
				break;
			}
			const char = this.take();
			switch (char) {
				case '(': {
					this.groupCount++;
					if (this.groupCount > MAX_GROUPS)
						throw new FindPatternError('A search pattern may hold nine groups');
					const inner = this.sequence(depth + 1);
					if (this.peek() !== ')') throw new FindPatternError('Unmatched ( in the search pattern');
					this.index++;
					atoms.push(`(${inner})`);
					break;
				}
				case '?':
					atoms.push(WILDCARD_ANY);
					break;
				case '*':
					atoms.push(`${WILDCARD_ANY}*?`);
					break;
				case '<':
					atoms.push(`(?<![${WORD_CHARACTERS}])`);
					break;
				case '>':
					atoms.push(`(?![${WORD_CHARACTERS}])`);
					break;
				case '[':
					atoms.push(this.characterSet());
					break;
				case '@':
					this.quantify(atoms, '+');
					break;
				case '{':
					this.quantify(atoms, this.repeatCount());
					break;
				case '\\':
					if (this.done) throw new FindPatternError('The search pattern ends with \\');
					atoms.push(escapeFindChar(this.take()));
					break;
				case '^':
					atoms.push(findCodeSource(this.take()));
					break;
				default:
					atoms.push(escapeFindChar(char));
			}
		}
		return atoms.join('');
	}

	/** Repeats the last atom: `a@` is `a+`, and `a{2,3}` is `a` two or three times. */
	private quantify(atoms: string[], quantifier: string): void {
		const last = atoms.pop();
		if (last === undefined) throw new FindPatternError('Nothing to repeat in the search pattern');
		atoms.push(`(?:${last})${quantifier}`);
	}

	/** Reads `n}`, `n,}` or `n,m}` (the opening brace is already read) as a quantifier. */
	private repeatCount(): string {
		const close = this.query.indexOf('}', this.index);
		if (close < 0) throw new FindPatternError('Unmatched { in the search pattern');
		const body = this.query.slice(this.index, close);
		this.index = close + 1;
		const match = /^(\d+)(?:(,)(\d*))?$/.exec(body);
		if (!match) throw new FindPatternError('Invalid repeat count in the search pattern');
		const min = Number(match[1]);
		if (match[2] === undefined) return `{${min}}`;
		if (match[3] === '') return `{${min},}`;
		const max = Number(match[3]);
		if (max < min) throw new FindPatternError('A repeat count has a maximum below its minimum');
		return `{${min},${max}}`;
	}

	/** Reads `[abc]`, `[a-z]` or `[!x]` (the opening bracket is already read). */
	private characterSet(): string {
		let negated = false;
		if (this.peek() === '!') {
			negated = true;
			this.index++;
		}
		const items: string[] = [];
		for (;;) {
			if (this.done) throw new FindPatternError('Unmatched [ in the search pattern');
			if (this.peek() === ']') {
				this.index++;
				break;
			}
			const first = this.take();
			const isRange =
				this.peek() === '-' &&
				this.index + 1 < this.query.length &&
				this.query.charAt(this.index + 1) !== ']';
			if (!isRange) {
				items.push(classCharacter(first));
				continue;
			}
			this.index++;
			const last = this.take();
			if ((first.codePointAt(0) ?? 0) > (last.codePointAt(0) ?? 0))
				throw new FindPatternError('A character range runs backwards in the search pattern');
			items.push(`${classCharacter(first)}-${classCharacter(last)}`);
		}
		if (!items.length) throw new FindPatternError('Empty [ ] in the search pattern');
		return `[${negated ? '^' : ''}${items.join('')}]`;
	}
}
