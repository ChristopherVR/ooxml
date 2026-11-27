import type { Paragraph, Table, TableCell, TextRun } from './model';

const MAX_CELL_TEXT = 4096;
const MAX_FORMULA_CELLS = 10_000;
const NUMBER =
	/^([+-]?)\s*(?:\p{Sc}\s*)?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(e[+-]?\d+)?\s*(?:\p{Sc}\s*)?(%)?$/iu;
const DIRECTIONS = new Map([
	['ABOVE', [-1, 0]],
	['BELOW', [1, 0]],
	['LEFT', [0, -1]],
	['RIGHT', [0, 1]],
]);

function cellText(cell: TableCell | undefined): string {
	let text = '';
	let runs = 0;
	for (const paragraph of cell?.paragraphs ?? []) {
		for (const run of paragraph.runs) {
			if (++runs > MAX_CELL_TEXT) return '';
			if (run.fieldChar || run.fieldCode !== undefined) continue;
			if (text.length + run.text.length > MAX_CELL_TEXT) return '';
			text += run.text;
		}
		text += ' ';
	}
	return text.trim();
}

const accountingText = (text: string): boolean => text.startsWith('(') && text.endsWith(')');

function parseTableNumber(text: string): number {
	const accounting = accountingText(text);
	const match = NUMBER.exec(accounting ? text.slice(1, -1).trim() : text);
	if (!match?.[2] || (accounting && match[1])) return NaN;
	const value = Number(`${match[1]}${match[2].replaceAll(',', '')}${match[3] ?? ''}`);
	return (accounting ? -value : value) / (match[4] ? 100 : 1);
}

/** Word's positional table functions. Reads cached results in referenced cells. */
export function evaluateTableFormula(
	table: Table,
	row: number,
	column: number,
	instruction: string,
): string | undefined {
	if (instruction.length > 128) return;
	const match =
		/^\s*=\s*(SUM|AVERAGE|COUNT|MAX|MIN|PRODUCT)\s*\(\s*([A-Z]+(?:\s*[,;]\s*[A-Z]+){0,3})\s*\)\s*$/i.exec(
			instruction,
		);
	if (!match?.[1] || !match[2]) return;
	const width = table.rows[row]?.length ?? 0;
	if (column < 0 || column >= width || table.rows.length * width > MAX_FORMULA_CELLS) return;
	const values: number[] = [];
	let firstValueText: string | undefined;
	const directions = new Set(match[2].toUpperCase().split(/\s*[,;]\s*/));
	for (const direction of directions) {
		const delta = DIRECTIONS.get(direction);
		const [dy, dx] = delta ?? [];
		if (dy === undefined || dx === undefined) return;
		let plainRun = false;
		for (
			let y = row + dy, x = column + dx;
			y >= 0 && y < table.rows.length && x >= 0 && x < width;
			y += dy, x += dx
		) {
			if (table.rowProperties?.[y]?.header) continue;
			const text = cellText(table.rows[y]?.[x]);
			const inner = accountingText(text) ? text.slice(1, -1).trim() : '';
			const signed = /^[+-]/.test(inner);
			const accounting = Boolean(inner) && !signed;
			const number = parseTableNumber(signed ? inner : text);
			// Word skips leading labels to locate the plain numeric range. Accounting values
			// before that range contribute, but an accounting value after it marks its end.
			if (plainRun && (accounting || !Number.isFinite(number))) break;
			if (!Number.isFinite(number)) continue;
			firstValueText ??= text;
			values.push(number);
			plainRun ||= !accounting;
		}
	}
	const sum = () => values.reduce((total, value) => total + value, 0);
	const operation = match[1].toUpperCase();
	const result =
		operation === 'COUNT'
			? values.length
			: operation === 'SUM'
				? sum()
				: operation === 'PRODUCT'
					? values.length
						? values.reduce((product, value) => product * value, 1)
						: 0
					: operation === 'AVERAGE'
						? sum() / values.length
						: operation === 'MIN'
							? values.length
								? Math.min(...values)
								: 0
							: values.length
								? Math.max(...values)
								: 0;
	if (!Number.isFinite(result)) return undefined;
	const rounded = Number(result.toPrecision(15));
	if (
		operation !== 'COUNT' &&
		firstValueText &&
		accountingText(firstValueText) &&
		!/^\(\s*[+-]/.test(firstValueText)
	) {
		const currency = firstValueText.match(/\p{Sc}/u)?.[0];
		const magnitude = currency
			? currency +
				Math.abs(rounded).toLocaleString('en-US', {
					minimumFractionDigits: 2,
					maximumFractionDigits: 2,
				})
			: String(Math.abs(rounded));
		return rounded < 0 ? `(${magnitude})` : magnitude;
	}
	return String(rounded);
}

/** A complex Word field with a cached result that Word can update. */
export function tableFormulaRuns(
	table: Table,
	row: number,
	column: number,
	instruction: string,
): TextRun[] | undefined {
	const text = evaluateTableFormula(table, row, column, instruction);
	if (text === undefined) return;
	const instr = instruction.trim();
	return [
		{ text: '', fieldChar: 'begin' },
		{ text: '', fieldCode: ` ${instr} ` },
		{ text: '', fieldChar: 'separate' },
		{ text, field: { instr } },
		{ text: '', fieldChar: 'end' },
	];
}

export interface TableSort {
	column: number;
	order: 'ascending' | 'descending';
	type: 'text' | 'number';
	headerRows: number;
	locale: string;
}

/** Stable sort. Complete rows, including their properties, move together. */
export function sortTableRows(table: Table, settings: TableSort): Table {
	const { column, order, type, headerRows, locale } = settings;
	const collator = new Intl.Collator(locale, { sensitivity: 'base' });
	const rows = table.rows.map((cells, index) => ({ cells, index, text: cellText(cells[column]) }));
	const body = rows.slice(headerRows).sort((left, right) => {
		if (type === 'text')
			return collator.compare(left.text, right.text) * (order === 'ascending' ? 1 : -1);
		const a = parseTableNumber(left.text);
		const b = parseTableNumber(right.text);
		if (!Number.isFinite(a)) return Number.isFinite(b) ? 1 : 0;
		if (!Number.isFinite(b)) return -1;
		return (a - b) * (order === 'ascending' ? 1 : -1);
	});
	const sorted = [...rows.slice(0, headerRows), ...body];
	return {
		...table,
		rows: sorted.map((row) => row.cells),
		...(table.rowProperties && {
			rowProperties: sorted.map((row) => table.rowProperties?.[row.index] ?? {}),
		}),
	};
}

export type TableTextSeparator = 'tabs' | 'commas' | 'paragraphs';

/** Converts simple table rows to paragraphs, preserving inline formatting and cell paragraph breaks. */
export function tableToText(table: Table, separator: TableTextSeparator): Paragraph[] {
	if (separator === 'paragraphs')
		return table.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs));
	return table.rows.flatMap((row) => {
		const output: Paragraph[] = [];
		for (const cell of row) {
			const first = cell.paragraphs[0];
			if (!first) continue;
			const previous = output.at(-1);
			if (previous) {
				previous.runs.push({ text: separator === 'tabs' ? '\t' : ',' });
				for (const run of first.runs) previous.runs.push(run);
			} else output.push({ ...first, runs: [...first.runs] });
			for (const paragraph of cell.paragraphs.slice(1))
				output.push({ ...paragraph, runs: [...paragraph.runs] });
		}
		return output;
	});
}
