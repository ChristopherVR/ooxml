// Rewriting structured (table) references as A1 references, as "Convert to Range" does.
import { columnLabel, quoteSheetName } from '../address.js';
import type { Table } from '../model.js';
import { FormulaError, type StructuredRef } from './ast.js';
import { parseStructured } from './structured.js';
import { joinTokens, type Token, tokenize } from './tokenizer.js';

export interface StructuredTarget {
	table: Table;
	/** Name of the sheet holding the table. */
	tableSheet: string;
	/** Sheet of the formula being rewritten (adds a sheet prefix when it differs). */
	formulaSheet: string;
	/** Row of the formula cell, for `[@Col]` (`#This Row`) references. */
	row?: number;
	/** The formula lies inside the table, so bare `[Col]` references mean this table. */
	inside?: boolean;
}

const abs = (row: number, col: number, rowAbs = true): string =>
	`$${columnLabel(col)}${rowAbs ? '$' : ''}${row + 1}`;

/** The A1 text of one structured reference, or `#REF!` when it cannot be resolved. */
function a1For(ref: StructuredRef, t: StructuredTarget): string {
	const { table } = t;
	const { start, end } = table.range;
	const index = (name: string | undefined): number | undefined => {
		if (name === undefined) return undefined;
		const i = table.columns.findIndex((c) => c.name.toLowerCase() === name.toLowerCase());
		return i < 0 ? -1 : start.col + i;
	};
	const c1 = index(ref.column);
	const c2 = index(ref.column2) ?? c1;
	if (c1 === -1 || c2 === -1) return '#REF!';
	const colLo = c1 ?? start.col;
	const colHi = c2 ?? end.col;
	const headerRow = table.headerRow ? start.row : undefined;
	const totalsRow = table.totalsRow ? end.row : undefined;
	const dataLo = start.row + (table.headerRow ? 1 : 0);
	const dataHi = end.row - (table.totalsRow ? 1 : 0);
	const specials = new Set(ref.specials);
	let rowLo: number;
	let rowHi: number;
	let thisRow = false;
	if (specials.has('#This Row')) {
		if (t.row === undefined) return '#REF!';
		rowLo = rowHi = t.row;
		thisRow = true;
	} else if (specials.has('#All')) {
		rowLo = start.row;
		rowHi = end.row;
	} else {
		const rows: number[] = [];
		const data = specials.has('#Data') || (!specials.has('#Headers') && !specials.has('#Totals'));
		if (specials.has('#Headers')) {
			if (headerRow === undefined) return '#REF!';
			rows.push(headerRow);
		}
		if (data) rows.push(dataLo, dataHi);
		if (specials.has('#Totals')) {
			if (totalsRow === undefined) return '#REF!';
			rows.push(totalsRow);
		}
		rowLo = Math.min(...rows);
		rowHi = Math.max(...rows);
	}
	const prefix =
		t.formulaSheet.toLowerCase() === t.tableSheet.toLowerCase()
			? ''
			: `${quoteSheetName(t.tableSheet)}!`;
	const a = abs(rowLo, colLo, !thisRow);
	if (rowLo === rowHi && colLo === colHi) return prefix + a;
	return `${prefix}${a}:${abs(rowHi, colHi, !thisRow)}`;
}

/**
 * Replaces references to `target.table` (structured references and the bare table name) with
 * absolute A1 references. Unparsable formulas are returned unchanged.
 */
export function structuredToA1(formula: string, target: StructuredTarget): string {
	let tokens: Token[];
	try {
		tokens = tokenize(formula);
	} catch (e) {
		if (e instanceof FormulaError) return formula;
		throw e;
	}
	const name = target.table.name.toLowerCase();
	let changed = false;
	for (const token of tokens) {
		const value = typeof token.value === 'string' ? token.value.toLowerCase() : undefined;
		if (token.kind === 'structured' && (value === name || (value === '' && target.inside))) {
			try {
				token.text = a1For(parseStructured(token.text, target.table.name), target);
			} catch (e) {
				if (!(e instanceof FormulaError)) throw e;
				token.text = '#REF!';
			}
			changed = true;
		} else if (token.kind === 'name' && !token.prefix && value === name) {
			token.text = a1For({ table: target.table.name, specials: [] }, target);
			changed = true;
		}
	}
	return changed ? joinTokens(tokens) : formula;
}
