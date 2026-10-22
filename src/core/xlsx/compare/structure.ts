// Defined names and the cheap structural parts of a sheet. Validations and conditional formats
// (see `conditional.ts`) have no identity of their own, so each is compared as a whole: a changed
// one is reported as removed and added.
import { formatRange, type CellRange } from '../address';
import type { DataValidation, DefinedName, Table, Workbook, Worksheet } from '../model';
import { stableKey } from '../stable-key';
import { conditionalFormatDifferences } from './conditional';
import type { SheetPair } from './sheets';
import type { CompareCategory, SheetDifference, WorkbookDifference } from './types';

const byPosition = (a: CellRange, b: CellRange) =>
	a.start.row - b.start.row ||
	a.start.col - b.start.col ||
	a.end.row - b.end.row ||
	a.end.col - b.end.col;

/** Items of `before` missing from `after` and the other way round, each in its own order. */
function setDifference<T>(
	before: readonly T[],
	after: readonly T[],
	key: (item: T) => string,
): { removed: T[]; added: T[] } {
	const beforeKeys = new Set(before.map(key));
	const afterKeys = new Set(after.map(key));
	return {
		removed: before.filter((item) => !afterKeys.has(key(item))),
		added: after.filter((item) => !beforeKeys.has(key(item))),
	};
}

/** A defined name and, for a sheet-scoped one, its sheet's name as a report shows it. */
interface ScopedName {
	name: DefinedName;
	/** Set only for a sheet-scoped name, so the rest of an entry spreads into a difference. */
	scope?: string;
}

/**
 * One workbook's defined names by lower-case name and scope. A name on a paired sheet is scoped by
 * the pair (and shown with the after sheet's name), so it matches across the two workbooks; a name
 * on an unpaired sheet has a scope of its own side, so it never does.
 */
function namesByKey(
	book: Workbook,
	side: 'before' | 'after',
	pairs: readonly SheetPair[],
	after: Workbook,
): Map<string, ScopedName> {
	const pairOf = new Map(pairs.map((pair, index) => [pair[side], index]));
	const out = new Map<string, ScopedName>();
	for (const name of book.definedNames) {
		const key = name.name.toLowerCase();
		const local = name.localSheet;
		if (local === undefined) {
			out.set(`${key}\n`, { name });
			continue;
		}
		const index = pairOf.get(local);
		const pair = index === undefined ? undefined : pairs[index];
		const scopeId = pair ? `p${index}` : `${side}${local}`;
		const sheet = pair ? after.sheets[pair.after] : book.sheets[local];
		out.set(`${key}\n${scopeId}`, { name, scope: sheet?.name ?? `#${local}` });
	}
	return out;
}

/** Names pair case-insensitively, as Excel matches them; a new spelling is still a change. */
const sameDefinition = (a: DefinedName, b: DefinedName): boolean =>
	a.name === b.name &&
	a.formula === b.formula &&
	!a.hidden === !b.hidden &&
	(a.comment ?? '') === (b.comment ?? '');

/**
 * Defined names added, removed or changed, sorted by name. A sheet-scoped name is matched through
 * the sheet pairing, so moving its sheet does not change it.
 */
export function nameDifferences(
	before: Workbook,
	after: Workbook,
	pairs: readonly SheetPair[],
): WorkbookDifference[] {
	const was = namesByKey(before, 'before', pairs, after);
	const now = namesByKey(after, 'after', pairs, after);
	const found = new Map<string, WorkbookDifference>();
	for (const [key, { name, ...scope }] of was) {
		const next = now.get(key)?.name;
		if (!next) found.set(key, { kind: 'nameRemoved', name: name.name, before: name, ...scope });
		else if (!sameDefinition(name, next))
			found.set(key, { kind: 'nameChanged', name: next.name, before: name, after: next, ...scope });
	}
	for (const [key, { name, ...scope }] of now)
		if (!was.has(key))
			found.set(key, { kind: 'nameAdded', name: name.name, after: name, ...scope });
	return [...found.keys()].sort().map((key) => found.get(key)!);
}

type TableField = Extract<SheetDifference, { kind: 'tableChanged' }>['fields'][number];

const column = (c: Table['columns'][number]) => ({
	name: c.name,
	totalsRowFunction: c.totalsRowFunction,
	totalsRowLabel: c.totalsRowLabel,
	calculatedColumnFormula: c.calculatedColumnFormula,
});

const tableStyle = (t: Table) => ({
	name: t.styleName,
	rows: !!t.showRowStripes,
	columns: !!t.showColumnStripes,
	first: !!t.showFirstColumn,
	last: !!t.showLastColumn,
});

function tableFields(a: Table, b: Table): TableField[] {
	const fields: TableField[] = [];
	if (formatRange(a.range) !== formatRange(b.range)) fields.push('range');
	if (stableKey(a.columns.map(column)) !== stableKey(b.columns.map(column))) fields.push('columns');
	if (a.headerRow !== b.headerRow) fields.push('headerRow');
	if (a.totalsRow !== b.totalsRow) fields.push('totalsRow');
	if (stableKey(tableStyle(a)) !== stableKey(tableStyle(b))) fields.push('style');
	return fields;
}

const tableName = (t: Table) => (t.displayName || t.name).toLowerCase();

/**
 * A validation without its range order and defaults: `between`, `stop`, flags that are off, empty
 * texts, and a drop-down that is shown (only `showDropDown: false` hides it).
 */
const validationKey = (v: DataValidation) =>
	stableKey({
		...v,
		ranges: v.ranges.map(formatRange).sort(),
		operator: v.operator === 'between' ? undefined : v.operator,
		allowBlank: v.allowBlank || undefined,
		showDropDown: v.showDropDown === false ? false : undefined,
		showErrorMessage: v.showErrorMessage || undefined,
		showInputMessage: v.showInputMessage || undefined,
		errorStyle: v.errorStyle === 'stop' ? undefined : v.errorStyle,
		errorTitle: v.errorTitle || undefined,
		error: v.error || undefined,
		promptTitle: v.promptTitle || undefined,
		prompt: v.prompt || undefined,
	});

/** Merges, tables, data validation and conditional formats of a sheet pair. */
export function sheetStructure(
	before: Worksheet,
	after: Worksheet,
	categories: ReadonlySet<CompareCategory>,
): SheetDifference[] {
	const out: SheetDifference[] = [];
	if (categories.has('merge')) {
		const { removed, added } = setDifference(before.merges, after.merges, formatRange);
		const merges = [
			...removed.map((range) => ({ kind: 'mergeRemoved' as const, range })),
			...added.map((range) => ({ kind: 'mergeAdded' as const, range })),
		];
		for (const merge of merges.sort((a, b) => byPosition(a.range, b.range))) out.push(merge);
	}
	if (categories.has('table')) {
		const now = new Map(after.tables.map((t) => [tableName(t), t]));
		const was = new Set(before.tables.map(tableName));
		for (const table of before.tables) {
			const next = now.get(tableName(table));
			if (!next) out.push({ kind: 'tableRemoved', table: table.displayName, before: table });
			else {
				const fields = tableFields(table, next);
				if (fields.length)
					out.push({
						kind: 'tableChanged',
						table: next.displayName,
						before: table,
						after: next,
						fields,
					});
			}
		}
		for (const table of after.tables)
			if (!was.has(tableName(table)))
				out.push({ kind: 'tableAdded', table: table.displayName, after: table });
	}
	if (categories.has('dataValidation')) {
		const { removed, added } = setDifference(
			before.dataValidations,
			after.dataValidations,
			validationKey,
		);
		for (const v of removed) out.push({ kind: 'dataValidationRemoved', before: v });
		for (const v of added) out.push({ kind: 'dataValidationAdded', after: v });
	}
	if (categories.has('conditionalFormat'))
		for (const d of conditionalFormatDifferences(
			before.conditionalFormats,
			after.conditionalFormats,
		))
			out.push(d);
	return out;
}
