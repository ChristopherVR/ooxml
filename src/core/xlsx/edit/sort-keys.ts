// Sort levels beyond plain values: cell and font colour, conditional-format icon, custom lists
// and case sensitivity. Colour and icon levels put the matching rows first (ascending, "On Top")
// or last (descending, "On Bottom") and leave the rest in their order.
import { cellFillColor, cellFontColor } from '../filter-criteria';
import type { ConditionalFormatResult } from '../layout/types';
import type { Cell, CellValue, SortBy, Workbook } from '../model';
import { isCellError } from '../model';

export interface SortKey {
	/** Column index, or row index when sorting left to right. */
	col: number;
	descending?: boolean;
	/** What the level orders by; values by default. */
	sortOn?: SortBy;
	/** `#RRGGBB` the colour levels match; absent means "No Cell Color" / automatic font colour. */
	color?: string;
	/** Icon the icon level matches; `index` absent means "No Cell Icon". */
	icon?: { set: string; index?: number };
	/** Custom list order for value levels. */
	customList?: readonly string[];
}

export interface SortOptions {
	/** Reorder columns by the values in rows instead of rows by the values in columns. */
	leftToRight?: boolean;
	/** Lower case sorts before upper case within otherwise equal text. */
	caseSensitive?: boolean;
	/** Needed for colour levels (styles and theme). */
	workbook?: Workbook;
	/** Conditional-format output of a cell, so colour and icon levels see it too. */
	formatAt?: (row: number, col: number) => ConditionalFormatResult | undefined;
}

/** Excel's ascending sort ranks: numbers, then text, then logicals, then errors. */
function rank(value: CellValue): number {
	if (typeof value === 'number') return 0;
	if (typeof value === 'string') return 1;
	if (typeof value === 'boolean') return 2;
	if (isCellError(value)) return 3;
	return 4;
}

const collator = new Intl.Collator(undefined, { sensitivity: 'base' });
const caseCollator = new Intl.Collator(undefined, { sensitivity: 'variant', caseFirst: 'lower' });

const isBlank = (value: CellValue): boolean => value === null || value === '';

export interface CompareOptions {
	caseSensitive?: boolean;
	customList?: readonly string[];
}

function listIndex(list: readonly string[], value: CellValue, caseSensitive: boolean): number {
	if (typeof value !== 'string') return -1;
	const text = value.trim();
	return list.findIndex((item) =>
		caseSensitive ? item === text : item.toLowerCase() === text.toLowerCase(),
	);
}

/**
 * Compares two cell values as Excel's sort does. Blanks always sort last, whatever the direction;
 * text compares case-insensitively unless `caseSensitive` (lower case first), and values found
 * in a custom list come first, in list order.
 */
export function compareCellValues(
	a: CellValue,
	b: CellValue,
	descending = false,
	options: CompareOptions = {},
): number {
	const blankA = isBlank(a);
	const blankB = isBlank(b);
	if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
	let result: number;
	const list = options.customList;
	const ia = list ? listIndex(list, a, !!options.caseSensitive) : -1;
	const ib = list ? listIndex(list, b, !!options.caseSensitive) : -1;
	if (ia >= 0 || ib >= 0) {
		// Listed values precede unlisted ones; the direction flips the list order only.
		if (ia < 0 || ib < 0) return ia >= 0 ? -1 : 1;
		result = ia - ib;
	} else {
		const ra = rank(a);
		const rb = rank(b);
		if (ra !== rb) result = ra - rb;
		else if (typeof a === 'number' && typeof b === 'number') result = a - b;
		else if (typeof a === 'string' && typeof b === 'string')
			result = collator.compare(a, b) || (options.caseSensitive ? caseCollator.compare(a, b) : 0);
		else if (typeof a === 'boolean' && typeof b === 'boolean') result = Number(a) - Number(b);
		else result = 0;
	}
	return descending ? -result : result;
}

/** What one level compares for one cell: the value, or whether it matches the colour or icon. */
export type Probe = { value: CellValue } | { match: boolean };

const normalColor = (color: string): string =>
	(color.startsWith('#') ? color : `#${color}`).toUpperCase();

/** Reads the comparison input of `key` for the cell at (`row`, `col`). */
export function probeCell(
	key: SortKey,
	cell: Cell | undefined,
	row: number,
	col: number,
	options: SortOptions,
): Probe {
	const on = key.sortOn ?? 'value';
	if (on === 'value') return { value: cell?.value ?? null };
	const cf = options.formatAt?.(row, col);
	if (on === 'icon') {
		const icon = cf?.icon;
		const want = key.icon;
		return {
			match:
				want?.index === undefined ? !icon : icon?.set === want.set && icon.index === want.index,
		};
	}
	const workbook = options.workbook;
	if (!workbook) return { match: false };
	const want = key.color ? normalColor(key.color) : undefined;
	if (on === 'cellColor') {
		const have = cellFillColor(workbook, cell, cf);
		return { match: want ? have === want : have === undefined };
	}
	return { match: cellFontColor(workbook, cell, cf) === (want ?? '#000000') };
}

/** Compares two probes of `key`. */
export function compareProbes(key: SortKey, a: Probe, b: Probe, options: SortOptions): number {
	if ('match' in a && 'match' in b) {
		const result = Number(b.match) - Number(a.match);
		return key.descending ? -result : result;
	}
	if ('value' in a && 'value' in b)
		return compareCellValues(a.value, b.value, key.descending, {
			caseSensitive: !!options.caseSensitive,
			...(key.customList ? { customList: key.customList } : {}),
		});
	return 0;
}
