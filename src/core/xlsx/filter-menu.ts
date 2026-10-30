// The model behind an AutoFilter drop-down: the value checklist, the colours on offer, the kind
// of the column and the criteria the menu's presets build. Pure, so the DOM only paints it.
import { getCell } from './cells';
import { compileAutoFilter } from './filter-criteria';
import { cellFillColor, cellFontColor } from './filter-criteria';
import { displayText } from './display-text';
import { isDateFormat } from './numfmt/format';
import type {
	AutoFilter,
	AutoFilterColumn,
	CustomFilterCondition,
	CustomFilters,
	FilterCriteria,
	FilterOperator,
	Workbook,
	Worksheet,
} from './model';
import { styleAt } from './styles';

export type FilterColumnKind = 'text' | 'number' | 'date';

export interface FilterMenuValue {
	text: string;
	checked: boolean;
	/** Date or number value of the first row showing the text. */
	number?: number;
}

export interface FilterMenuModel {
	kind: FilterColumnKind;
	/** Distinct non-blank display texts of the rows the other columns leave visible. */
	values: FilterMenuValue[];
	/** Present when the candidate rows hold blank cells. */
	blanks?: { checked: boolean };
	/** `#RRGGBB` fills of the candidate rows; `noFill` when some have none. */
	fills: string[];
	noFill: boolean;
	fonts: string[];
	hasCriteria: boolean;
}

/** Rows scanned for the checklist (Excel lists every row; the menu stays responsive). */
const MAX_ROWS = 100_000;

/** Whether a column has any criteria (its drop-down shows the funnel). */
export const hasFilterCriteria = (column: AutoFilterColumn | undefined): boolean =>
	!!column &&
	!!(
		column.values ||
		column.blank ||
		column.dateGroups?.length ||
		column.custom ||
		column.top10 ||
		column.dynamic ||
		column.color ||
		column.icon
	);

/** Number, date or text, by the majority of the column's non-blank cells. */
export function filterColumnKind(
	workbook: Workbook,
	sheet: Worksheet,
	filter: AutoFilter,
	offset: number,
): FilterColumnKind {
	const col = filter.range.start.col + offset;
	let numbers = 0;
	let dates = 0;
	let texts = 0;
	const end = Math.min(filter.range.end.row, filter.range.start.row + 2000);
	for (let row = filter.range.start.row + 1; row <= end; row++) {
		const cell = getCell(sheet, row, col);
		if (typeof cell?.value === 'number') {
			numbers++;
			if (isDateFormat(styleAt(workbook, cell.styleId).numFmt)) dates++;
		} else if (cell?.value !== null && cell?.value !== undefined && cell.value !== '') texts++;
	}
	if (numbers === 0 || texts > numbers) return 'text';
	return dates * 2 > numbers ? 'date' : 'number';
}

export function filterMenuModel(
	workbook: Workbook,
	sheet: Worksheet,
	filter: AutoFilter,
	offset: number,
): FilterMenuModel {
	const column = filter.columns?.find((c) => c.offset === offset);
	const col = filter.range.start.col + offset;
	const others = compileAutoFilter(workbook, sheet, {
		range: filter.range,
		columns: (filter.columns ?? []).filter((c) => c.offset !== offset),
	});
	const own = compileAutoFilter(workbook, sheet, {
		range: filter.range,
		columns: column ? [column] : [],
	});
	const seen = new Map<string, FilterMenuValue>();
	const fills = new Set<string>();
	const fonts = new Set<string>();
	let blanks: { checked: boolean } | undefined;
	let noFill = false;
	const last = Math.min(filter.range.end.row, filter.range.start.row + MAX_ROWS);
	for (let row = filter.range.start.row + 1; row <= last; row++) {
		if (!others(row)) continue;
		const cell = getCell(sheet, row, col);
		const checked = own(row);
		const text = displayText(workbook, cell);
		if (text === '') blanks = { checked: (blanks?.checked ?? false) || checked };
		else {
			const known = seen.get(text);
			if (known) known.checked ||= checked;
			else
				seen.set(text, {
					text,
					checked,
					...(typeof cell?.value === 'number' ? { number: cell.value } : {}),
				});
		}
		const fill = cellFillColor(workbook, cell);
		if (fill) fills.add(fill);
		else noFill = true;
		fonts.add(cellFontColor(workbook, cell));
	}
	const values = [...seen.values()].sort((a, b) =>
		a.number !== undefined && b.number !== undefined
			? a.number - b.number
			: a.number !== undefined
				? -1
				: b.number !== undefined
					? 1
					: a.text.localeCompare(b.text, undefined, { sensitivity: 'base' }),
	);
	return {
		kind: filterColumnKind(workbook, sheet, filter, offset),
		values,
		...(blanks ? { blanks } : {}),
		fills: [...fills].sort(),
		noFill,
		fonts: [...fonts].sort(),
		hasCriteria: hasFilterCriteria(column),
	};
}

/** The criteria for a checklist choice; `undefined` (clear the filter) when everything is ticked. */
export function checklistCriteria(
	model: FilterMenuModel,
	checked: ReadonlySet<string>,
	blanksChecked: boolean,
): FilterCriteria | undefined {
	const all = model.values.every((v) => checked.has(v.text)) && (!model.blanks || blanksChecked);
	if (all) return undefined;
	return {
		values: model.values.filter((v) => checked.has(v.text)).map((v) => v.text),
		...(blanksChecked && model.blanks ? { blank: true as const } : {}),
	};
}

// --- Presets of the Text, Number and Date filter submenus ------------------------------------

/** The comparison operators of the Custom AutoFilter dialog. */
export type CustomOperator =
	| 'equals'
	| 'notEquals'
	| 'greaterThan'
	| 'greaterThanOrEqual'
	| 'lessThan'
	| 'lessThanOrEqual'
	| 'beginsWith'
	| 'notBeginsWith'
	| 'endsWith'
	| 'notEndsWith'
	| 'contains'
	| 'notContains';

const FILE_OPERATOR: Record<string, FilterOperator> = {
	equals: 'equal',
	notEquals: 'notEqual',
	greaterThan: 'greaterThan',
	greaterThanOrEqual: 'greaterThanOrEqual',
	lessThan: 'lessThan',
	lessThanOrEqual: 'lessThanOrEqual',
};

/** Escapes `* ? ~` so they match themselves inside a wildcard pattern. */
const escapeWildcards = (text: string): string => text.replace(/[~*?]/g, '~$&');

/** One dialog row as a stored condition (wildcards carry begins / ends / contains). */
export function encodeCondition(op: CustomOperator, value: string): CustomFilterCondition {
	const direct = FILE_OPERATOR[op];
	if (direct) return { operator: direct, value };
	const safe = escapeWildcards(value);
	const negated = op.startsWith('not');
	const kind = op.replace(/^not/, '').replace(/^./, (c) => c.toLowerCase());
	const pattern =
		kind === 'beginsWith' ? `${safe}*` : kind === 'endsWith' ? `*${safe}` : `*${safe}*`;
	return { operator: negated ? 'notEqual' : 'equal', value: pattern };
}

/** The dialog row a stored condition shows as; wildcard patterns that fit no preset stay `equals`. */
export function decodeCondition(condition: CustomFilterCondition): {
	op: CustomOperator;
	value: string;
} {
	const { operator, value } = condition;
	const direct = (Object.keys(FILE_OPERATOR) as CustomOperator[]).find(
		(key) => FILE_OPERATOR[key] === operator,
	);
	if ((operator === 'equal' || operator === 'notEqual') && /[*?]/.test(value.replace(/~./g, ''))) {
		const not = operator === 'notEqual';
		const unescape = (s: string) => s.replace(/~(.)/g, '$1');
		const inner = value.slice(1, -1);
		const plain = (s: string) => !/[*?]/.test(s.replace(/~./g, ''));
		if (value.length > 2 && value.startsWith('*') && value.endsWith('*') && plain(inner))
			return { op: not ? 'notContains' : 'contains', value: unescape(inner) };
		if (value.endsWith('*') && plain(value.slice(0, -1)))
			return { op: not ? 'notBeginsWith' : 'beginsWith', value: unescape(value.slice(0, -1)) };
		if (value.startsWith('*') && plain(value.slice(1)))
			return { op: not ? 'notEndsWith' : 'endsWith', value: unescape(value.slice(1)) };
	}
	return { op: direct ?? 'equals', value };
}

/** A custom filter from one or two dialog rows. */
export function customFilter(
	rows: ReadonlyArray<{ op: CustomOperator; value: string }>,
	and: boolean,
): FilterCriteria {
	const conditions = rows.slice(0, 2).map((r) => encodeCondition(r.op, r.value));
	const custom: CustomFilters = {
		...(and && conditions.length > 1 ? { and: true } : {}),
		conditions,
	};
	return { custom };
}

/** Between: greater than or equal to `low` AND less than or equal to `high`. */
export const betweenFilter = (low: string, high: string): FilterCriteria =>
	customFilter(
		[
			{ op: 'greaterThanOrEqual', value: low },
			{ op: 'lessThanOrEqual', value: high },
		],
		true,
	);

/** A fill or font colour filter for a `#RRGGBB` colour (`undefined`: no fill / automatic). */
export function colorFilterCriteria(color: string | undefined, cellColor: boolean): FilterCriteria {
	const rgb = color && `FF${color.replace('#', '').toUpperCase()}`;
	const style = cellColor
		? {
				fill: rgb
					? { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { rgb } }
					: { type: 'pattern' as const, pattern: 'none' as const },
			}
		: rgb
			? { font: { color: { rgb } } }
			: {};
	return { color: { cellColor, style } };
}

export const topFilterCriteria = (
	top: boolean,
	percent: boolean,
	value: number,
): FilterCriteria => ({
	top10: { top, percent, value },
});
