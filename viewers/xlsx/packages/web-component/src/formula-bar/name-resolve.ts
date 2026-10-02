// Pure Name Box logic: parse what was typed into a target (a reference or a defined name), and
// find the defined name that matches the current selection.
import {
	columnLabel,
	normalizeRange,
	parseRange,
	quoteSheetName,
	sheetByName,
	validateDefinedName,
	type CellRange,
	type DefinedName,
	type Workbook,
} from '@christophervr/xlsx-core';

export interface NameTarget {
	sheet: number;
	range: CellRange;
}

export type NameBoxResult =
	| ({ kind: 'go' } & NameTarget)
	| { kind: 'define'; name: string }
	| { kind: 'invalid' };

/** Parses `A1`, `B2:C5`, `A:C`, `1:3`, `Sheet2!A1`, `'My sheet'!$A$1:$B$2`. */
export function parseSheetReference(
	workbook: Workbook,
	activeSheet: number,
	text: string,
): NameTarget | undefined {
	const ref = text.trim().replace(/^=/, '');
	const bang = ref.lastIndexOf('!');
	let sheet = activeSheet;
	if (bang >= 0) {
		const raw = ref.slice(0, bang);
		const name = /^'.*'$/.test(raw) ? raw.slice(1, -1).replace(/''/g, "'") : raw;
		const target = sheetByName(workbook, name);
		if (!target) return undefined;
		sheet = workbook.sheets.indexOf(target);
	}
	const range = parseRange(ref.slice(bang + 1).replace(/\$/g, ''));
	return range ? { sheet, range: normalizeRange(range) } : undefined;
}

/** The defined name visible from `sheet` (sheet-scoped names win over workbook names). */
export function findDefinedName(
	workbook: Workbook,
	sheet: number,
	name: string,
): DefinedName | undefined {
	const lower = name.toLowerCase();
	const matches = workbook.definedNames.filter((d) => d.name.toLowerCase() === lower);
	return (
		matches.find((d) => d.localSheet === sheet) ?? matches.find((d) => d.localSheet === undefined)
	);
}

export function resolveNameBox(
	workbook: Workbook,
	activeSheet: number,
	text: string,
): NameBoxResult {
	const input = text.trim();
	if (!input) return { kind: 'invalid' };
	const ref = parseSheetReference(workbook, activeSheet, input);
	if (ref) return { kind: 'go', ...ref };
	const named = findDefinedName(workbook, activeSheet, input);
	if (named) {
		const target = parseSheetReference(workbook, named.localSheet ?? activeSheet, named.formula);
		return target ? { kind: 'go', ...target } : { kind: 'invalid' };
	}
	return validateDefinedName(input) === undefined
		? { kind: 'define', name: input }
		: { kind: 'invalid' };
}

const absoluteCell = (row: number, col: number): string => `$${columnLabel(col)}$${row + 1}`;

/** `Sheet1!$A$1:$B$2` (one cell: `Sheet1!$A$1`). */
export function absoluteReference(sheetName: string, range: CellRange): string {
	const { start, end } = normalizeRange(range);
	const a = absoluteCell(start.row, start.col);
	const tail =
		start.row === end.row && start.col === end.col ? '' : `:${absoluteCell(end.row, end.col)}`;
	return `${quoteSheetName(sheetName)}!${a}${tail}`;
}

const sameRange = (a: CellRange, b: CellRange): boolean =>
	a.start.row === b.start.row &&
	a.start.col === b.start.col &&
	a.end.row === b.end.row &&
	a.end.col === b.end.col;

/** A defined name whose reference is exactly this range on this sheet. */
export function nameForRange(
	workbook: Workbook,
	sheet: number,
	range: CellRange,
): string | undefined {
	const want = normalizeRange(range);
	for (const named of workbook.definedNames) {
		if (named.hidden || (named.localSheet !== undefined && named.localSheet !== sheet)) continue;
		const target = parseSheetReference(workbook, named.localSheet ?? sheet, named.formula);
		if (target && target.sheet === sheet && sameRange(target.range, want)) return named.name;
	}
	return undefined;
}
