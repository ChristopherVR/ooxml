/** Quotes a TSV field the way Excel does when it holds a tab, line break or quote. */
function quoteField(text: string): string {
	return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/** Tab-separated text for the system clipboard (Excel style: CRLF rows, trailing CRLF). */
export function toTsv(rows: string[][]): string {
	if (!rows.length) return '';
	return rows.map((row) => row.map(quoteField).join('\t')).join('\r\n') + '\r\n';
}

/**
 * Parses tab-separated clipboard text. Quoted fields may hold tabs, line breaks and doubled
 * quotes; one trailing line break is ignored. Rows are padded to the same width.
 */
export function parseTsv(text: string): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let field = '';
	let i = 0;
	let atFieldStart = true;
	const body = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
	if (body === '') return text === '' ? [] : [['']];
	while (i < body.length) {
		const ch = body[i] ?? '';
		if (atFieldStart && ch === '"') {
			const close = findClosingQuote(body, i + 1);
			if (close >= 0) {
				const next = body[close + 1];
				if (next === undefined || next === '\t' || next === '\n') {
					field = body.slice(i + 1, close).replace(/""/g, '"');
					i = close + 1;
					atFieldStart = false;
					continue;
				}
			}
		}
		atFieldStart = false;
		if (ch === '\t') {
			row.push(field);
			field = '';
			atFieldStart = true;
		} else if (ch === '\n') {
			row.push(field);
			rows.push(row);
			row = [];
			field = '';
			atFieldStart = true;
		} else field += ch;
		i++;
	}
	row.push(field);
	rows.push(row);
	const width = Math.max(...rows.map((r) => r.length));
	for (const r of rows) while (r.length < width) r.push('');
	return rows;
}

/** Index of the quote closing a quoted field that opened before `from` (`""` is an escape). */
function findClosingQuote(text: string, from: number): number {
	for (let i = from; i < text.length; i++) {
		if (text[i] !== '"') continue;
		if (text[i + 1] === '"') {
			i++;
			continue;
		}
		return i;
	}
	return -1;
}
