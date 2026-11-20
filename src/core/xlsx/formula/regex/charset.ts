/** The inline flags in force where an atom appears: the pattern's, changed by `(?ims-ims:...)`. */
export interface Flags {
	readonly i: boolean;
	readonly m: boolean;
	readonly s: boolean;
}

/**
 * Whether one code point matches one atom: a literal, `.`, a class escape such as `\d` or `\p{L}`,
 * or a bracketed class. A literal without `i` is compared directly. Every other atom is decided by
 * the platform's RegExp on that single code point, so classes, Unicode properties and case folding
 * mean exactly what they mean in JavaScript; a test that reads one code point cannot backtrack.
 * Answers are cached, ASCII in a table.
 */
export class CharMatcher {
	/** For each ASCII code point, 1 when it matches. */
	readonly ascii = new Uint8Array(128);
	private readonly native: RegExp | null;
	private readonly other = new Map<number, boolean>();

	/** `literal` is the code point of a case-sensitive literal, or -1 for `source` in `flags`. */
	constructor(
		readonly literal: number,
		source: string,
		flags: Flags,
	) {
		this.native =
			literal >= 0
				? null
				: new RegExp(`^(?:${source})$`, `u${flags.i ? 'i' : ''}${flags.s ? 's' : ''}`);
		for (let cp = 0; cp < 128; cp++) this.ascii[cp] = this.ask(cp) ? 1 : 0;
	}

	test(cp: number): boolean {
		if (cp < 128) return this.ascii[cp] === 1;
		let hit = this.other.get(cp);
		if (hit === undefined) {
			hit = this.ask(cp);
			if (this.other.size < 65_536) this.other.set(cp, hit);
		}
		return hit;
	}

	private ask(cp: number): boolean {
		return this.native ? this.native.test(String.fromCodePoint(cp)) : cp === this.literal;
	}
}

/** A literal code point: compared directly, or through case folding under `i`. */
export function literalMatcher(cp: number, flags: Flags): CharMatcher {
	return flags.i
		? new CharMatcher(-1, `\\u{${cp.toString(16)}}`, flags)
		: new CharMatcher(cp, '', flags);
}

const HIGH_START = 0xd800;
const LOW_START = 0xdc00;
const LOW_END = 0xdfff;

/** The code point that starts at `at` (a surrogate pair counts as one). */
export function codePointAt(text: string, at: number): number {
	const c = text.charCodeAt(at);
	if (c >= HIGH_START && c < LOW_START && at + 1 < text.length) {
		const d = text.charCodeAt(at + 1);
		if (d >= LOW_START && d <= LOW_END) return (c - HIGH_START) * 0x400 + d - LOW_START + 0x10000;
	}
	return c;
}

/** The code point that ends at `at`. */
export function codePointBefore(text: string, at: number): number {
	const c = text.charCodeAt(at - 1);
	if (c >= LOW_START && c <= LOW_END && at >= 2) {
		const d = text.charCodeAt(at - 2);
		if (d >= HIGH_START && d < LOW_START) return (d - HIGH_START) * 0x400 + c - LOW_START + 0x10000;
	}
	return c;
}

/** Code units of a code point. */
export const width = (cp: number): number => (cp > 0xffff ? 2 : 1);

export const isLineTerminator = (unit: number): boolean =>
	unit === 10 || unit === 13 || unit === 0x2028 || unit === 0x2029;
