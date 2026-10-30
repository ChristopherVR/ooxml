// Error Checking: the background rules Excel marks with a green triangle, as an ordered list of
// issues a dialog can walk through (Next / Previous / Ignore Error).
import { type CellRange, rangeContains } from '../address';
import { forEachCell } from '../cells';
import { translateFormula } from '../formula/transform';
import { parseNumberText } from '../formula/text-number';
import { isCellError, type Workbook } from '../model';
import { styleAt } from '../styles';
import { type CellRef, formulaReferences } from './trace';

export type ErrorRule =
	| 'error-value'
	| 'inconsistent-formula'
	| 'omits-adjacent-cells'
	| 'number-as-text'
	| 'unlocked-formula';

export const ERROR_RULES: readonly ErrorRule[] = [
	'error-value',
	'number-as-text',
	'inconsistent-formula',
	'omits-adjacent-cells',
	'unlocked-formula',
];

export interface ErrorIssue extends CellRef {
	rule: ErrorRule;
	/** The error shown (`#DIV/0!`) for `error-value`. */
	code?: string;
}

const issueKey = (i: CellRef & { rule: ErrorRule }): string =>
	`${i.sheet}:${i.row}:${i.col}:${i.rule}`;

const ignoredByBook = new WeakMap<Workbook, Set<string>>();
const ignoredOf = (workbook: Workbook): Set<string> => {
	let set = ignoredByBook.get(workbook);
	if (!set) ignoredByBook.set(workbook, (set = new Set()));
	return set;
};

/** Marks an issue as ignored (Ignore Error); it stays hidden until {@link resetIgnoredErrors}. */
export function ignoreError(workbook: Workbook, issue: ErrorIssue): void {
	ignoredOf(workbook).add(issueKey(issue));
}

export function resetIgnoredErrors(workbook: Workbook): void {
	ignoredByBook.delete(workbook);
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

/** Every issue on a sheet (all sheets when `sheet` is undefined), in reading order, minus ignored ones. */
export function findErrors(
	workbook: Workbook,
	sheet?: number,
	rules: readonly ErrorRule[] = ERROR_RULES,
): ErrorIssue[] {
	const on = new Set(rules);
	const ignored = ignoredOf(workbook);
	const out: ErrorIssue[] = [];
	workbook.sheets.forEach((ws, index) => {
		if (sheet !== undefined && sheet !== index) return;
		const found: ErrorIssue[] = [];
		forEachCell(ws, (cell, row, col) => {
			const at = { sheet: index, row, col };
			const add = (rule: ErrorRule, extra?: { code: string }): void => {
				const issue: ErrorIssue = { ...at, rule, ...extra };
				if (!ignored.has(issueKey(issue))) found.push(issue);
			};
			if (cell.formula) {
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
