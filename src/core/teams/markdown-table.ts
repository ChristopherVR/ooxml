export interface MarkdownTable {
	headers: string[];
	align: ('left' | 'center' | 'right' | null)[];
	rows: string[][];
}

/** Split unescaped pipes, including escaped pipes inside inline code. */
function cells(raw: string): { values: string[]; hasPipe: boolean } {
	const line = raw.trim();
	const values: string[] = [];
	let current = '';
	let lastPipe = -1;
	for (let i = 0; i < line.length; i++) {
		const char = line[i]!;
		if (char === '\\' && i + 1 < line.length) {
			const next = line[++i]!;
			current += next === '|' ? '|' : `\\${next}`;
		} else if (char === '|') {
			values.push(current.trim());
			current = '';
			lastPipe = i;
		} else current += char;
	}
	values.push(current.trim());
	if (line.startsWith('|')) values.shift();
	if (lastPipe === line.length - 1) values.pop();
	return { values, hasPipe: lastPipe >= 0 };
}

/** Top-level GFM-style tables only; nested container parsing is intentionally unsupported. */
export function markdownTable(
	lines: string[],
	start: number,
): { table: MarkdownTable; end: number } | null {
	const header = cells(lines[start] ?? '');
	const delimiter = cells(lines[start + 1] ?? '');
	if (
		!(header.hasPipe || delimiter.hasPipe) ||
		!header.values.length ||
		header.values.length > 128 ||
		header.values.length !== delimiter.values.length ||
		delimiter.values.some((value) => !/^:?-+:?$/u.test(value))
	)
		return null;
	const table: MarkdownTable = {
		headers: header.values,
		align: delimiter.values.map((value) =>
			value.startsWith(':')
				? value.endsWith(':')
					? 'center'
					: 'left'
				: value.endsWith(':')
					? 'right'
					: null,
		),
		rows: [],
	};
	let end = start + 1;
	for (let i = start + 2; i < lines.length; i++) {
		// Bound sparse-table expansion. Remaining source lines stay visible as ordinary blocks.
		if ((table.rows.length + 2) * table.headers.length > 16_384) break;
		const line = lines[i]!;
		if (
			!line.trim() ||
			/^\s{0,3}(?:>|#{1,6}\s|`{3,}|~{3,}|[-*+]\s|\d+[.)]\s|(?:-\s*){3,}$|(?:\*\s*){3,}$|(?:_\s*){3,}$)/u.test(
				line,
			)
		)
			break;
		const row = cells(line).values;
		table.rows.push(table.headers.map((_, column) => row[column] ?? ''));
		end = i;
	}
	return { table, end };
}
