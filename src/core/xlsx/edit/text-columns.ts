// Text to Columns: the pure splitting rules the dialog previews and applies. Delimited splitting
// honours a text qualifier; fixed-width splitting cuts each line at character positions.

export interface SplitOptions {
	delimiters: string[];
	consecutive: boolean;
	qualifier: string;
}

/** Splits one line at any of the delimiters, honouring a text qualifier (`"a,b",c`). */
export function splitDelimited(line: string, options: SplitOptions): string[] {
	const out: string[] = [];
	let current = '';
	let quoted = false;
	for (let i = 0; i < line.length; i++) {
		const ch = line[i] ?? '';
		if (options.qualifier && ch === options.qualifier) {
			if (quoted && line[i + 1] === options.qualifier) {
				current += ch;
				i++;
			} else quoted = !quoted;
			continue;
		}
		if (!quoted && options.delimiters.includes(ch)) {
			out.push(current);
			current = '';
			if (options.consecutive) while (options.delimiters.includes(line[i + 1] ?? '\u0000')) i++;
			continue;
		}
		current += ch;
	}
	out.push(current);
	return out;
}

/**
 * Splits one line at fixed character positions. A break at position `p` ends a column after the
 * `p`th character (1-based), so breaks `[5]` split `ABCDEFGHIJ` into `ABCDE` and `FGHIJ`. Every
 * break gives a column, so a line shorter than a break ends in empty columns, the same count for
 * every line. Characters are code points, so a surrogate pair is never cut.
 */
export function splitFixedWidth(line: string, breaks: readonly number[]): string[] {
	const chars = Array.from(line);
	const cuts = [...new Set(breaks)].sort((a, b) => a - b);
	const out: string[] = [];
	let start = 0;
	for (const cut of cuts) {
		out.push(chars.slice(start, cut).join(''));
		start = cut;
	}
	out.push(chars.slice(start).join(''));
	return out;
}

/**
 * Parses the break positions typed in the dialog (`5, 12`): whole numbers from 1, returned
 * ascending and unique. An empty entry means no breaks. Undefined when any entry is not such a
 * number.
 */
export function parseBreakPositions(text: string): number[] | undefined {
	const parts = text
		.split(',')
		.map((part) => part.trim())
		.filter((part) => part !== '');
	const values: number[] = [];
	for (const part of parts) {
		if (!/^\d+$/.test(part)) return undefined;
		const value = Number(part);
		if (!Number.isSafeInteger(value) || value < 1) return undefined;
		values.push(value);
	}
	return [...new Set(values)].sort((a, b) => a - b);
}
