// Resolves reference text (chart series, list sources) to the values it covers.
import { type CellRange, normalizeRange } from '../address.js';
import { getValue, usedRange } from '../cells.js';
import { type CalcEngine, type Token, tokenize } from '../formula/index.js';
import type { CellValue, Workbook } from '../model.js';
import { sheetByName } from '../workbook.js';
import type { EvaluateRef } from './chart-view.js';

/** Largest number of cells a direct read returns (whole columns are clipped to the used area). */
const MAX_CELLS = 1_000_000;

export interface RefEvaluatorOptions {
	/** Sheet unqualified references point at (default: the active sheet). */
	sheet?: number;
}

/** The sheet and range of a single plain A1 reference, or undefined for anything else. */
function plainReference(
	workbook: Workbook,
	text: string,
	sheetIndex: number,
): { sheet: number; range: CellRange } | undefined {
	let tokens: Token[];
	try {
		tokens = tokenize(text).filter((t) => t.kind !== 'ws');
	} catch {
		return undefined;
	}
	const token = tokens[0];
	if (tokens.length !== 1 || !token || token.kind !== 'ref' || !token.ref) return undefined;
	const prefix = token.prefix;
	if (prefix && (prefix.book !== undefined || prefix.sheet2 !== undefined)) return undefined;
	const named = prefix ? sheetByName(workbook, prefix.sheet) : undefined;
	const sheet = prefix ? (named ? workbook.sheets.indexOf(named) : -1) : sheetIndex;
	if (sheet < 0 || !workbook.sheets[sheet]) return undefined;
	const { start, end } = token.ref;
	return {
		sheet,
		range: normalizeRange({
			start: { row: start.row, col: start.col },
			end: { row: end.row, col: end.col },
		}),
	};
}

function readRange(workbook: Workbook, sheetIndex: number, range: CellRange): CellValue[] {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) return [];
	const used = usedRange(sheet);
	if (!used) return [];
	const end = {
		row: Math.min(range.end.row, Math.max(range.start.row, used.end.row)),
		col: Math.min(range.end.col, Math.max(range.start.col, used.end.col)),
	};
	if ((end.row - range.start.row + 1) * (end.col - range.start.col + 1) > MAX_CELLS) return [];
	const out: CellValue[] = [];
	for (let r = range.start.row; r <= end.row; r++)
		for (let c = range.start.col; c <= end.col; c++) out.push(getValue(sheet, r, c));
	return out;
}

/**
 * An `evaluateRef` for {@link chartView} (and any other "values of this reference" need):
 * `Sheet1!$B$2:$B$9`, `'My Sheet'!A1:A5` or a bare `A1:A5` read row by row from the live cells.
 * With a calc engine, anything else (defined names, `OFFSET(...)`, unions) is evaluated by
 * `calc.evaluateArray` and flattened row-major; without one, such references give `[]`.
 * Whole-column references stop at the sheet's used area.
 */
export function createRefEvaluator(
	workbook: Workbook,
	calc?: CalcEngine,
	options: RefEvaluatorOptions = {},
): EvaluateRef {
	return (ref: string): CellValue[] => {
		const text = ref.trim().replace(/^=/, '');
		if (!text) return [];
		const sheetIndex = options.sheet ?? workbook.activeSheet;
		const plain = plainReference(workbook, text, sheetIndex);
		if (plain) return readRange(workbook, plain.sheet, plain.range);
		if (!calc) return [];
		try {
			return calc.evaluateArray(text, { sheet: sheetIndex, row: 0, col: 0 }).flat();
		} catch {
			return [];
		}
	};
}
