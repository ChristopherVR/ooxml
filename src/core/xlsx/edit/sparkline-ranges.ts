import { MAX_COL, MAX_ROW, formatRange, normalizeRange, type CellRange } from '../address';
import type { Sparkline, Workbook } from '../model';
import { parseFormula } from '../formula/parser';

export function checkedSparklineRange(raw: CellRange): CellRange {
	const range = normalizeRange(raw);
	for (const point of [range.start, range.end])
		if (
			!Number.isInteger(point.row) ||
			!Number.isInteger(point.col) ||
			point.row < 0 ||
			point.col < 0 ||
			point.row > MAX_ROW ||
			point.col > MAX_COL
		)
			throw new RangeError('The reference is not valid.');
	return range;
}

/** One contiguous local or sheet-qualified range, without external workbooks or 3D references. */
export function sparklineDataRange(
	workbook: Workbook,
	sheet: number,
	source: string,
): { sheet: number; range: CellRange } {
	let ast;
	try {
		ast = parseFormula(source.trim().replace(/^=/, ''));
	} catch {
		throw new Error('The reference is not valid.');
	}
	if (
		ast.type !== 'ref' ||
		!ast.ref ||
		!['cell', 'area'].includes(ast.ref.kind) ||
		ast.prefix?.book ||
		ast.prefix?.sheet2 ||
		ast.spill
	)
		throw new Error('The reference is not valid.');
	const index = ast.prefix
		? workbook.sheets.findIndex((ws) => ws.name.toLowerCase() === ast.prefix!.sheet.toLowerCase())
		: sheet;
	if (!workbook.sheets[index]) throw new Error('The reference is not valid.');
	return { sheet: index, range: checkedSparklineRange({ start: ast.ref.start, end: ast.ref.end }) };
}

/** Maps a one-column destination to data rows, or a one-row destination to data columns. */
export function newSparklines(
	workbook: Workbook,
	sheet: number,
	source: string,
	locations: CellRange,
): Sparkline[] {
	const data = sparklineDataRange(workbook, sheet, source);
	const dest = checkedSparklineRange(locations);
	const rows = dest.end.row - dest.start.row + 1;
	const cols = dest.end.col - dest.start.col + 1;
	const dr = data.range.end.row - data.range.start.row + 1;
	const dc = data.range.end.col - data.range.start.col + 1;
	if (
		(rows > 1 && cols > 1) ||
		(rows === 1 && cols === 1 && dr > 1 && dc > 1) ||
		(rows > 1 && rows !== dr) ||
		(cols > 1 && cols !== dc)
	)
		throw new Error('The data range and location range do not match.');
	if (Math.max(rows, cols) > 100_000) throw new Error('The sparkline location range is too large.');
	const prefix = `'${workbook.sheets[data.sheet]!.name.replaceAll("'", "''")}'!`;
	const out: Sparkline[] = [];
	for (let i = 0; i < Math.max(rows, cols); i++) {
		const range = structuredClone(data.range);
		if (rows > 1) range.start.row = range.end.row = data.range.start.row + i;
		if (cols > 1) range.start.col = range.end.col = data.range.start.col + i;
		const host = {
			row: dest.start.row + (rows > 1 ? i : 0),
			col: dest.start.col + (cols > 1 ? i : 0),
		};
		out.push({ formula: prefix + formatRange(range), host: { start: host, end: { ...host } } });
	}
	return out;
}
