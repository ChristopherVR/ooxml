// Evaluation of one AutoFilter column's criteria (ECMA-376 18.3.2): value checklist with date
// groups, custom conditions, top 10, dynamic, colour and icon filters.
import { getCell } from './cells';
import { displayText } from './display-text';
import { inDateGroup, inDynamicPeriod, todaySerial } from './filter-dynamic';
import { resolveColor } from './layout/colors';
import type { ConditionalFormatResult } from './layout/types';
import type {
	AutoFilter,
	AutoFilterColumn,
	Cell,
	CustomFilterCondition,
	Fill,
	Workbook,
	Worksheet,
} from './model';
import { styleAt } from './styles';

export interface FilterEvalOptions {
	/** Today's date serial for date filters; defaults to the local clock. */
	today?: number;
	/** Conditional-format result of a cell, so colour and icon filters see CF output too. */
	formatAt?: (row: number, col: number) => ConditionalFormatResult | undefined;
}

type RowTest = (row: number) => boolean;

const numeric = (cell: Cell | undefined): number | undefined =>
	typeof cell?.value === 'number' ? cell.value : undefined;

function wildcardRegex(pattern: string): RegExp | undefined {
	let out = '';
	let wild = false;
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i]!;
		if (ch === '~' && i + 1 < pattern.length) {
			wild = true;
			out += pattern[++i]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		} else if (ch === '*') {
			out += '[\\s\\S]*';
			wild = true;
		} else if (ch === '?') {
			out += '[\\s\\S]';
			wild = true;
		} else out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	}
	return wild ? new RegExp(`^${out}$`, 'i') : undefined;
}

const rank = (v: Cell['value']): number =>
	typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;

function compare(cell: Cell | undefined, text: string, operand: string): number {
	const value = cell?.value ?? null;
	const asNumber = operand.trim() === '' ? NaN : Number(operand);
	if (typeof value === 'number' && !Number.isNaN(asNumber)) return value - asNumber;
	if (typeof value === 'string') {
		const a = value.toLowerCase();
		const b = operand.toLowerCase();
		return a < b ? -1 : a > b ? 1 : 0;
	}
	// A number against text, or a boolean: numbers sort before text before booleans.
	return rank(value) - (Number.isNaN(asNumber) ? 1 : 0);
}

function conditionTest(
	workbook: Workbook,
	condition: CustomFilterCondition,
): (cell: Cell | undefined) => boolean {
	const { operator, value: operand } = condition;
	if (operand === ' ' && (operator === 'notEqual' || operator === 'equal')) {
		const blank = (cell: Cell | undefined) =>
			cell?.value === null || cell?.value === undefined || cell.value === '';
		return operator === 'notEqual' ? (cell) => !blank(cell) : blank;
	}
	const wild = operator === 'equal' || operator === 'notEqual' ? wildcardRegex(operand) : undefined;
	return (cell) => {
		const empty = cell?.value === null || cell?.value === undefined || cell.value === '';
		if (empty)
			return operator === 'notEqual' ? operand !== '' : operator === 'equal' && operand === '';
		if (wild) {
			const hit = wild.test(displayText(workbook, cell));
			return operator === 'equal' ? hit : !hit;
		}
		const cmp = compare(cell, displayText(workbook, cell), operand);
		switch (operator) {
			case 'equal':
				return cmp === 0;
			case 'notEqual':
				return cmp !== 0;
			case 'lessThan':
				return cmp < 0;
			case 'lessThanOrEqual':
				return cmp <= 0;
			case 'greaterThan':
				return cmp > 0;
			case 'greaterThanOrEqual':
				return cmp >= 0;
		}
	};
}

export const SOLID = (fill: Fill | undefined) =>
	fill?.type === 'pattern' && fill.pattern !== 'none' ? (fill.fgColor ?? fill.bgColor) : undefined;

/** The `#RRGGBB` fill of a cell (conditional format included), or `undefined` for no fill. */
export function cellFillColor(
	workbook: Workbook,
	cell: Cell | undefined,
	cf?: ConditionalFormatResult,
): string | undefined {
	const fromCf = SOLID(cf?.style?.fill);
	if (fromCf) return resolveColor(fromCf, workbook.theme)?.toUpperCase();
	if (cf?.colorScale) return cf.colorScale.toUpperCase();
	const fill = styleAt(workbook, cell?.styleId).fill;
	return resolveColor(SOLID(fill), workbook.theme)?.toUpperCase();
}

/** The `#RRGGBB` font colour of a cell (conditional format included); automatic is black. */
export function cellFontColor(
	workbook: Workbook,
	cell: Cell | undefined,
	cf?: ConditionalFormatResult,
): string {
	const color = cf?.style?.font?.color ?? styleAt(workbook, cell?.styleId).font.color;
	return (resolveColor(color, workbook.theme, '#000000') ?? '#000000').toUpperCase();
}

function colourTest(
	workbook: Workbook,
	column: AutoFilterColumn,
	read: (row: number) => { cell: Cell | undefined; cf: ConditionalFormatResult | undefined },
): RowTest | undefined {
	const color = column.color;
	if (!color?.style) return undefined;
	if (color.cellColor) {
		const want = resolveColor(SOLID(color.style.fill), workbook.theme)?.toUpperCase();
		return (row) => {
			const { cell, cf } = read(row);
			return cellFillColor(workbook, cell, cf) === want;
		};
	}
	const want = (
		resolveColor(color.style.font?.color, workbook.theme, '#000000') ?? '#000000'
	).toUpperCase();
	return (row) => {
		const { cell, cf } = read(row);
		return cellFontColor(workbook, cell, cf) === want;
	};
}

/**
 * Builds the visibility test for the data rows of `filter`. Column statistics (top 10 cut-offs,
 * averages) are computed once, so build the test once per pass rather than per row.
 */
export function compileAutoFilter(
	workbook: Workbook,
	sheet: Worksheet,
	filter: AutoFilter,
	options: FilterEvalOptions = {},
): RowTest {
	const { range } = filter;
	const tests: RowTest[] = [];
	const today = options.today ?? todaySerial(workbook.date1904);
	for (const column of filter.columns ?? []) {
		const col = range.start.col + column.offset;
		const cellOf = (row: number) => getCell(sheet, row, col);
		const numbers = (): number[] => {
			const out: number[] = [];
			for (let r = range.start.row + 1; r <= range.end.row; r++) {
				const n = numeric(cellOf(r));
				if (n !== undefined) out.push(n);
			}
			return out;
		};
		if (column.top10) {
			const spec = column.top10;
			const sorted = numbers().sort((a, b) => (spec.top ? b - a : a - b));
			const count = spec.percent
				? Math.max(1, Math.floor((sorted.length * spec.value) / 100))
				: Math.max(1, Math.floor(spec.value));
			const cut = sorted[Math.min(count, sorted.length) - 1];
			tests.push((row) => {
				const n = numeric(cellOf(row));
				return n !== undefined && cut !== undefined && (spec.top ? n >= cut : n <= cut);
			});
		} else if (column.custom?.conditions.length) {
			const parts = column.custom.conditions.slice(0, 2).map((c) => conditionTest(workbook, c));
			const and = !!column.custom.and;
			tests.push((row) => {
				const cell = cellOf(row);
				return and ? parts.every((t) => t(cell)) : parts.some((t) => t(cell));
			});
		} else if (column.dynamic) {
			const { type } = column.dynamic;
			if (type === 'aboveAverage' || type === 'belowAverage') {
				const all = numbers();
				const avg = all.length ? all.reduce((a, b) => a + b, 0) / all.length : 0;
				tests.push((row) => {
					const n = numeric(cellOf(row));
					return n !== undefined && (type === 'aboveAverage' ? n > avg : n < avg);
				});
			} else if (type !== 'null')
				tests.push((row) => {
					const n = numeric(cellOf(row));
					return n !== undefined && inDynamicPeriod(type, n, today, workbook.date1904);
				});
		} else if (column.color) {
			const test = colourTest(workbook, column, (row) => ({
				cell: cellOf(row),
				cf: options.formatAt?.(row, col),
			}));
			if (test) tests.push(test);
		} else if (column.icon) {
			const { iconSet, iconId } = column.icon;
			if (options.formatAt)
				tests.push((row) => {
					const icon = options.formatAt?.(row, col)?.icon;
					return iconId === undefined ? !icon : icon?.set === iconSet && icon.index === iconId;
				});
		} else if (column.values || column.blank || column.dateGroups?.length) {
			const values = new Set(column.values ?? []);
			const groups = column.dateGroups ?? [];
			tests.push((row) => {
				const cell = cellOf(row);
				const text = displayText(workbook, cell);
				if (text === '') return !!column.blank;
				if (values.has(text)) return true;
				const n = numeric(cell);
				return n !== undefined && groups.some((g) => inDateGroup(g, n, workbook.date1904));
			});
		}
	}
	return (row) => tests.every((test) => test(row));
}

/** Whether a data row passes every column criterion of the filter. */
export function rowMatchesFilter(
	workbook: Workbook,
	sheet: Worksheet,
	row: number,
	filter: AutoFilter,
	options?: FilterEvalOptions,
): boolean {
	return compileAutoFilter(workbook, sheet, filter, options)(row);
}
