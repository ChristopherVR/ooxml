// Error Checking: the background rules Excel marks with a green triangle, as an ordered list of
// issues a dialog can walk through (Next / Previous / Ignore Error).
import { type CellRange, rangeContains } from '../address';
import { forEachCell } from '../cells';
import { translateFormula } from '../formula/transform';
import { parseNumberText } from '../formula/text-number';
import { isCellError, type IgnoredError, type IgnoredErrorFlag, type Workbook } from '../model';
import { styleAt } from '../styles';
import { type CellRef, formulaReferences } from './trace';

export type ErrorRule =
	| 'error-value'
	| 'inconsistent-formula'
	| 'omits-adjacent-cells'
	| 'number-as-text'
	| 'unlocked-formula'
	| 'empty-cell-reference'
	| 'calculated-column';

export const ERROR_RULES: readonly ErrorRule[] = [
	'error-value',
	'calculated-column',
	'number-as-text',
	'inconsistent-formula',
	'omits-adjacent-cells',
	'unlocked-formula',
	'empty-cell-reference',
];

/** The rules Excel has on by default (references to empty cells is off). */
export const DEFAULT_ERROR_RULES: readonly ErrorRule[] = ERROR_RULES.filter(
	(rule) => rule !== 'empty-cell-reference',
);

/** The  flag that switches each rule off for a range. */
export const RULE_FLAG: Readonly<Record<ErrorRule, IgnoredErrorFlag>> = {
	'error-value': 'evalError',
	'calculated-column': 'calculatedColumn',
	'number-as-text': 'numberStoredAsText',
	'inconsistent-formula': 'formula',
	'omits-adjacent-cells': 'formulaRange',
	'unlocked-formula': 'unlockedFormula',
	'empty-cell-reference': 'emptyCellReference',
};

export interface ErrorIssue extends CellRef {
	rule: ErrorRule;
	/** The error shown (#DIV/0!) for error-value. */
	code?: string;
}

const isIgnored = (workbook: Workbook, issue: ErrorIssue): boolean =>
	(workbook.sheets[issue.sheet]?.ignoredErrors ?? []).some(
		(e) => e.flags.includes(RULE_FLAG[issue.rule]) && e.ranges.some((r) => rangeContains(r, issue)),
	);

/** The sheet's ignored-error list with the issue added (Ignore Error), for `setIgnoredErrors`. */
export function withIgnoredError(workbook: Workbook, issue: ErrorIssue): IgnoredError[] {
	const list = structuredClone(workbook.sheets[issue.sheet]?.ignoredErrors ?? []);
	const flag = RULE_FLAG[issue.rule];
	const range = {
		start: { row: issue.row, col: issue.col },
		end: { row: issue.row, col: issue.col },
	};
	const same = list.find((e) => e.flags.length === 1 && e.flags[0] === flag);
	if (same) same.ranges.push(range);
	else list.push({ ranges: [range], flags: [flag] });
	return list;
}

const normal = (formula: string): string => formula.replace(/\s+/g, '').toUpperCase();

function inconsistent(workbook: Workbook, at: CellRef, formula: string): boolean {
	const rows = workbook.sheets[at.sheet]?.rows;
	const read = (dRow: number, dCol: number): string | undefined => {
		const f = rows?.get(at.row + dRow)?.get(at.col + dCol)?.formula;
		return f === undefined ? undefined : normal(translateFormula(f, -dRow, -dCol));
	};
	const mine = normal(formula);
	for (const [a, b] of [
		[read(0, -1), read(0, 1)],
		[read(-1, 0), read(1, 0)],
	] as const)
		if (a !== undefined && a === b && a !== mine) return true;
	return false;
}

function omitsAdjacent(workbook: Workbook, at: CellRef, formula: string): boolean {
	const ws = workbook.sheets[at.sheet];
	if (!ws) return false;
	const refs = formulaReferences(workbook, formula, at.sheet).filter((r) => r.sheet === at.sheet);
	const covered = (row: number, col: number): boolean =>
		refs.some((r) => rangeContains(r.range, { row, col }));
	const numeric = (row: number, col: number): boolean => {
		if (row < 0 || col < 0 || (row === at.row && col === at.col)) return false;
		return typeof ws.rows.get(row)?.get(col)?.value === 'number' && !covered(row, col);
	};
	return refs.some(({ range }: { range: CellRange }) => {
		const { start, end } = range;
		if (start.row === end.row && start.col === end.col) return false;
		if (start.col === end.col)
			return numeric(start.row - 1, start.col) || numeric(end.row + 1, start.col);
		if (start.row === end.row)
			return numeric(start.row, start.col - 1) || numeric(start.row, end.col + 1);
		return false;
	});
}

/**
 * Cells of calculated table columns whose formula differs from the rest of the column: the
 * column's formula is the one most data rows share (at least two, and more than half).
 */
function calculatedColumnOutliers(workbook: Workbook, sheet: number): Set<string> {
	const out = new Set<string>();
	const ws = workbook.sheets[sheet];
	for (const table of ws?.tables ?? []) {
		const first = table.range.start.row + (table.headerRow ? 1 : 0);
		const last = table.range.end.row - (table.totalsRow ? 1 : 0);
		if (!ws || last < first) continue;
		for (let col = table.range.start.col; col <= table.range.end.col; col++) {
			const forms = new Map<number, string>();
			const counts = new Map<string, number>();
			for (let row = first; row <= last; row++) {
				const f = ws.rows.get(row)?.get(col)?.formula;
				if (f === undefined) continue;
				const key = normal(translateFormula(f, first - row, 0));
				forms.set(row, key);
				counts.set(key, (counts.get(key) ?? 0) + 1);
			}
			const [mode, count] = [...counts].sort((a, b) => b[1] - a[1])[0] ?? ['', 0];
			if (count < 2 || count * 2 <= last - first + 1) continue;
			for (let row = first; row <= last; row++) {
				const cell = ws.rows.get(row)?.get(col);
				if (cell && (forms.get(row) ?? undefined) !== mode && (cell.formula || cell.value !== null))
					out.add(`${row},${col}`);
			}
		}
	}
	return out;
}

/** Whether a formula reads a single cell that holds nothing. */
function readsEmptyCell(workbook: Workbook, at: CellRef, formula: string): boolean {
	return formulaReferences(workbook, formula, at.sheet).some(({ sheet, range }) => {
		if (range.start.row !== range.end.row || range.start.col !== range.end.col) return false;
		const cell = workbook.sheets[sheet]?.rows.get(range.start.row)?.get(range.start.col);
		return !cell || (cell.value === null && !cell.formula);
	});
}

/**
 * Every issue on a sheet (all sheets when `sheet` is undefined), in reading order, minus the
 * ones the sheet's `ignoredErrors` switch off.
 */
export function findErrors(
	workbook: Workbook,
	sheet?: number,
	rules: readonly ErrorRule[] = DEFAULT_ERROR_RULES,
): ErrorIssue[] {
	const on = new Set(rules);
	const out: ErrorIssue[] = [];
	workbook.sheets.forEach((ws, index) => {
		if (sheet !== undefined && sheet !== index) return;
		const found: ErrorIssue[] = [];
		const outliers = on.has('calculated-column')
			? calculatedColumnOutliers(workbook, index)
			: new Set<string>();
		forEachCell(ws, (cell, row, col) => {
			const at = { sheet: index, row, col };
			const add = (rule: ErrorRule, extra?: { code: string }): void => {
				const issue: ErrorIssue = { ...at, rule, ...extra };
				if (!isIgnored(workbook, issue)) found.push(issue);
			};
			if (outliers.has(`${row},${col}`)) add('calculated-column');
			if (cell.formula) {
				if (on.has('empty-cell-reference') && readsEmptyCell(workbook, at, cell.formula))
					add('empty-cell-reference');
				if (on.has('error-value') && isCellError(cell.value))
					add('error-value', { code: cell.value.error });
				if (on.has('inconsistent-formula') && inconsistent(workbook, at, cell.formula))
					add('inconsistent-formula');
				if (on.has('omits-adjacent-cells') && omitsAdjacent(workbook, at, cell.formula))
					add('omits-adjacent-cells');
				if (
					on.has('unlocked-formula') &&
					styleAt(workbook, cell.styleId).protection?.locked === false
				)
					add('unlocked-formula');
			} else if (
				on.has('number-as-text') &&
				typeof cell.value === 'string' &&
				cell.value.trim() !== '' &&
				parseNumberText(cell.value) !== undefined
			)
				add('number-as-text');
		});
		out.push(...found);
	});
	return out;
}

/** The first issue after (or, going back, before) a cell, wrapping around; undefined when none. */
export function adjacentIssue(
	issues: readonly ErrorIssue[],
	from: CellRef,
	direction: 1 | -1,
	includeCurrent = false,
): ErrorIssue | undefined {
	if (!issues.length) return undefined;
	const order = (i: CellRef): number[] => [i.sheet, i.row, i.col];
	const cmp = (a: CellRef, b: CellRef): number => {
		const x = order(a);
		const y = order(b);
		for (let k = 0; k < 3; k++) if (x[k] !== y[k]) return (x[k] ?? 0) - (y[k] ?? 0);
		return 0;
	};
	const ahead = (i: ErrorIssue): boolean =>
		direction === 1
			? includeCurrent
				? cmp(i, from) >= 0
				: cmp(i, from) > 0
			: includeCurrent
				? cmp(i, from) <= 0
				: cmp(i, from) < 0;
	const candidates = issues.filter(ahead);
	if (candidates.length) return direction === 1 ? candidates[0] : candidates[candidates.length - 1];
	return direction === 1 ? issues[0] : issues[issues.length - 1];
}
