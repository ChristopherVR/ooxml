import { cellKey, rangeContains } from '../address.js';
import { forEachCellInRange, getValue } from '../cells.js';
import { translateFormula } from '../formula/index.js';
import { inTimePeriod } from '../time-period.js';
import { iconIndex, literal, localTodaySerial, mergeDxf, scaleColor } from './cf-helpers.js';

export { iconIndex, scaleColor } from './cf-helpers.js';
import type {
	CellValue,
	CfvoThreshold,
	ConditionalFormat,
	ConditionalRule,
	Workbook,
	Worksheet,
} from '../model.js';
import { isCellError } from '../model.js';
import { resolveColor } from './colors.js';
import {
	computeStats,
	isTruthy,
	percentile,
	testOperator,
	valueKey,
	valueText,
	type RangeStats,
} from './cf-values.js';
import type { ConditionalFormatEvaluator, ConditionalFormatResult } from './types.js';

export type CellAt = { sheet: number; row: number; col: number };
export type FormulaEvaluator = (formula: string, at: CellAt) => CellValue;

export interface ConditionalFormatOptions {
	/** Moves a formula's relative references by `dRow`/`dCol` (defaults to the formula module's). */
	translate?: (formula: string, dRow: number, dCol: number) => string;
	/**
	 * Today's date serial for "A Date Occurring" (`timePeriod`) rules. Defaults to evaluating
	 * `TODAY()` through the evaluator, then to the local clock.
	 */
	today?: () => number;
}

const GRAPHIC_RULES = new Set(['colorScale', 'dataBar', 'iconSet']);

interface Entry {
	format: ConditionalFormat;
	rule: ConditionalRule;
	order: number;
}

/**
 * Evaluates a sheet's conditional formats. Per-format statistics (min, max, average, ranks,
 * duplicates) are computed once and cached; create a new evaluator after the sheet changes.
 * Formulas are relative to the top-left cell of the format's first range, as in SpreadsheetML.
 */
export function createConditionalFormatEvaluator(
	workbook: Workbook,
	sheetIndex: number,
	evaluate: FormulaEvaluator,
	options: ConditionalFormatOptions = {},
): ConditionalFormatEvaluator {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) return { at: () => undefined };
	const translate = options.translate ?? translateFormula;
	const entries: Entry[] = [];
	let order = 0;
	for (const format of sheet.conditionalFormats)
		for (const rule of format.rules) entries.push({ format, rule, order: order++ });
	entries.sort((a, b) => a.rule.priority - b.rule.priority || a.order - b.order);

	const statsCache = new Map<ConditionalFormat, RangeStats>();
	const stats = (format: ConditionalFormat): RangeStats => {
		let cached = statsCache.get(format);
		if (!cached) {
			cached = computeStats(formatValues(sheet, format));
			statsCache.set(format, cached);
		}
		return cached;
	};
	const thresholdCache = new Map<Entry, number[]>();
	let todaySerial: number | undefined;
	const today = (): number => {
		if (todaySerial !== undefined) return todaySerial;
		let serial = options.today?.();
		if (serial === undefined) {
			const anchor = { sheet: sheetIndex, row: 0, col: 0 };
			let raw: CellValue = null;
			try {
				raw = evaluate('TODAY()', anchor);
			} catch {
				raw = null;
			}
			serial = typeof raw === 'number' ? raw : localTodaySerial(workbook.date1904);
		}
		todaySerial = serial;
		return serial;
	};

	const run = (formula: string, format: ConditionalFormat, row: number, col: number): CellValue => {
		const fixed = literal(formula);
		if (fixed !== undefined) return fixed;
		const anchor = format.ranges[0]?.start ?? { row, col };
		const dRow = row - anchor.row;
		const dCol = col - anchor.col;
		try {
			const moved = dRow || dCol ? translate(formula, dRow, dCol) : formula;
			return evaluate(moved, { sheet: sheetIndex, row, col });
		} catch {
			return { error: '#VALUE!' };
		}
	};

	const thresholdValue = (t: CfvoThreshold, entry: Entry): number => {
		const s = stats(entry.format);
		const anchor = entry.format.ranges[0]?.start ?? { row: 0, col: 0 };
		const n = (): number => {
			const raw = t.value === undefined ? 0 : run(t.value, entry.format, anchor.row, anchor.col);
			return typeof raw === 'number' ? raw : Number(raw) || 0;
		};
		switch (t.type) {
			case 'min':
				return s.min;
			case 'max':
				return s.max;
			case 'num':
			case 'formula':
				return n();
			case 'percent':
				return s.min + ((s.max - s.min) * n()) / 100;
			case 'percentile':
				return percentile(s.sorted, n() / 100);
		}
	};
	const thresholds = (entry: Entry, list: CfvoThreshold[]): number[] => {
		let cached = thresholdCache.get(entry);
		if (!cached) {
			cached = list.map((t) => thresholdValue(t, entry));
			thresholdCache.set(entry, cached);
		}
		return cached;
	};

	const matches = (entry: Entry, value: CellValue, row: number, col: number): boolean => {
		const { rule, format } = entry;
		switch (rule.type) {
			case 'cellIs': {
				const a = rule.formulas[0] !== undefined ? run(rule.formulas[0], format, row, col) : null;
				const b = rule.formulas[1] !== undefined ? run(rule.formulas[1], format, row, col) : null;
				return testOperator(rule.operator, value, a, b);
			}
			case 'expression':
				return isTruthy(run(rule.formula, format, row, col));
			case 'top10': {
				if (typeof value !== 'number') return false;
				const sorted = stats(format).sorted;
				if (!sorted.length) return false;
				const n = rule.percent
					? Math.max(1, Math.floor((sorted.length * rule.rank) / 100))
					: Math.max(1, rule.rank);
				const k = Math.min(n, sorted.length);
				return rule.bottom
					? value <= (sorted[k - 1] ?? Infinity)
					: value >= (sorted[sorted.length - k] ?? -Infinity);
			}
			case 'aboveAverage': {
				if (typeof value !== 'number') return false;
				const avg = stats(format).average;
				if (rule.below) return rule.equalAverage ? value <= avg : value < avg;
				return rule.equalAverage ? value >= avg : value > avg;
			}
			case 'duplicateValues':
			case 'uniqueValues': {
				const key = valueKey(value);
				if (!key) return false;
				const count = stats(format).counts.get(key) ?? 0;
				return rule.type === 'duplicateValues' ? count > 1 : count === 1;
			}
			case 'containsBlanks':
			case 'notContainsBlanks': {
				const blank = value === null || (typeof value === 'string' && value.trim() === '');
				return rule.type === 'containsBlanks' ? blank : !blank;
			}
			case 'containsErrors':
				return isCellError(value);
			case 'notContainsErrors':
				return !isCellError(value);
			case 'containsText':
			case 'notContainsText':
			case 'beginsWith':
			case 'endsWith': {
				if (isCellError(value)) return false;
				const text = valueText(value).toLowerCase();
				const needle = rule.text.toLowerCase();
				if (rule.type === 'containsText') return text.includes(needle);
				if (rule.type === 'notContainsText') return !text.includes(needle);
				if (rule.type === 'beginsWith') return text.startsWith(needle);
				return text.endsWith(needle);
			}
			case 'timePeriod':
				return (
					typeof value === 'number' &&
					inTimePeriod(rule.timePeriod, value, today(), workbook.date1904)
				);
			default:
				return false;
		}
	};

	/** Applies one rule to `result`; false when it does not apply to this cell. */
	const apply = (
		entry: Entry,
		value: CellValue,
		row: number,
		col: number,
		result: ConditionalFormatResult,
	): boolean => {
		const { rule } = entry;
		if (rule.type === 'colorScale') {
			if (result.colorScale || typeof value !== 'number') return false;
			const color = scaleColor(
				value,
				thresholds(entry, rule.thresholds),
				rule.colors.map((c) => resolveColor(c, workbook.theme, '#FFFFFF') ?? '#FFFFFF'),
			);
			if (!color) return false;
			result.colorScale = color;
			return true;
		}
		if (rule.type === 'dataBar') {
			if (result.dataBar || typeof value !== 'number') return false;
			const [lo = 0, hi = 0] = thresholds(entry, [rule.min, rule.max]);
			const fraction = hi > lo ? Math.max(0, Math.min(1, (value - lo) / (hi - lo))) : 1;
			result.dataBar = {
				fraction,
				color: resolveColor(rule.color, workbook.theme, '#638EC6') ?? '#638EC6',
			};
			if (value < 0) result.dataBar.negative = true;
			if (rule.showValue === false) result.hideValue = true;
			return true;
		}
		if (rule.type === 'iconSet') {
			if (result.icon || typeof value !== 'number') return false;
			const index = iconIndex(value, thresholds(entry, rule.thresholds), rule.thresholds);
			result.icon = {
				set: rule.iconSet,
				index: rule.reverse ? rule.thresholds.length - 1 - index : index,
			};
			if (rule.showValue === false) result.hideValue = true;
			return true;
		}
		if (!matches(entry, value, row, col)) return false;
		result.style = mergeDxf(result.style, rule.style);
		return true;
	};

	return {
		at(row, col) {
			const result: ConditionalFormatResult = {};
			let value: CellValue | undefined;
			let any = false;
			for (const entry of entries) {
				if (!entry.format.ranges.some((r) => rangeContains(r, { row, col }))) continue;
				value ??= getValue(sheet, row, col);
				if (!apply(entry, value, row, col, result)) continue;
				any = true;
				// "Stop If True" on a matching rule holds back every lower-priority rule. Excel 16
				// ignores it on colour scales, data bars and icon sets even when the file sets it.
				if (entry.rule.stopIfTrue && !GRAPHIC_RULES.has(entry.rule.type)) break;
			}
			return any ? result : undefined;
		},
	};
}

/** The values of every distinct stored cell inside a format's ranges. */
function formatValues(sheet: Worksheet, format: ConditionalFormat): CellValue[] {
	const seen = new Set<number>();
	const values: CellValue[] = [];
	for (const range of format.ranges)
		forEachCellInRange(sheet, range, (cell, row, col) => {
			const key = cellKey(row, col);
			if (seen.has(key)) return;
			seen.add(key);
			values.push(cell.value);
		});
	return values;
}
