/**
 * SVG path data (`d`) split into commands and their numbers, shared by the
 * path flattener and the path-to-segments reader.
 *
 * Numbers may run together (`10-5.5.5` is 10, -5.5, .5), and an arc's two
 * flags may be written without separators (`a1 1 0 011 1`), as the SVG path
 * grammar allows.
 *
 * @module geometry/svg-path-tokens
 */

/** One command letter and the numbers that follow it. */
export interface SvgPathToken {
	type: string;
	values: number[];
}

const COMMANDS = 'MLHVCSQTAZ';
const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/u;

/**
 * Split `d` into commands. A lenient read (the default) skips what is not a
 * command or a number; a strict read throws a `SyntaxError` for it and for
 * numbers before the first command.
 */
export function tokenizeSvgPath(d: string, options?: { strict?: boolean }): SvgPathToken[] {
	const strict = options?.strict === true;
	const tokens: SvgPathToken[] = [];
	let current: SvgPathToken | undefined;
	let i = 0;
	while (i < d.length) {
		const ch = d.charAt(i);
		if (ch === ',' || /\s/u.test(ch)) {
			i += 1;
			continue;
		}
		if (COMMANDS.includes(ch.toUpperCase()) && /[a-z]/iu.test(ch)) {
			current = { type: ch, values: [] };
			tokens.push(current);
			i += 1;
			continue;
		}
		const position = current ? current.values.length % 7 : -1;
		if (current?.type.toUpperCase() === 'A' && (position === 3 || position === 4)) {
			if (ch === '0' || ch === '1') {
				current.values.push(Number(ch));
				i += 1;
				continue;
			}
		}
		const match = NUMBER_RE.exec(d.slice(i));
		if (!match || !current) {
			if (strict) {
				throw new SyntaxError(
					current
						? `Unexpected "${ch}" in SVG path at ${i}`
						: 'SVG path data must start with a command',
				);
			}
			i += match ? match[0].length : 1;
			continue;
		}
		current.values.push(Number.parseFloat(match[0]));
		i += match[0].length;
	}
	return tokens;
}
