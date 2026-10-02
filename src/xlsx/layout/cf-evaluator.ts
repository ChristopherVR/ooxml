import { cellKey, rangeContains } from '../address.js';
import { forEachCellInRange, getValue } from '../cells.js';
import { translateFormula } from '../formula/index.js';
import { inTimePeriod } from '../time-period.js';
import { literal, localTodaySerial, mergeDxf } from './cf-helpers.js';
import type {
	CellValue,
	CfvoThreshold,
	ConditionalFormat,
	ConditionalRule,
	Workbook,
	Worksheet,
} from '../model.js';
import { isCellError } from '../model.js';
import { mixColors, resolveColor } from './colors.js';
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

	return {
		at(row, col) {
			const result: ConditionalFormatResult = {};
			let value: CellValue | undefined;
			let any = false;
			for (const entry of entries) {
				if (!entry.format.ranges.some((r) => rangeContains(r, { row, col }))) continue;
				value ??= getValue(sheet, row, col);
				const { rule } = entry;
				if (rule.type === 'colorScale') {
					if (result.colorScale || typeof value !== 'number') continue;
					const color = scaleColor(
						value,
						thresholds(entry, rule.thresholds),
						rule.colors.map((c) => resolveColor(c, workbook.theme, '#FFFFFF') ?? '#FFFFFF'),
					);
					if (color) {
						result.colorScale = color;
						any = true;
					}
					continue;
				}
				if (rule.type === 'dataBar') {
					if (result.dataBar || typeof value !== 'number') continue;
					const [lo = 0, hi = 0] = thresholds(entry, [rule.min, rule.max]);
					const fraction = hi > lo ? Math.max(0, Math.min(1, (value - lo) / (hi - lo))) : 1;
					result.dataBar = {
						fraction,
						color: resolveColor(rule.color, workbook.theme, '#638EC6') ?? '#638EC6',
					};
					if (value < 0) result.dataBar.negative = true;
					if (rule.showValue === false) result.hideValue = true;
					any = true;
					continue;
				}
				if (rule.type === 'iconSet') {
					if (result.icon || typeof value !== 'number') continue;
					const cuts = thresholds(entry, rule.thresholds);
					let index = 0;
					for (let i = 1; i < cuts.length; i++) if (value >= (cuts[i] ?? Infinity)) index = i;
					if (rule.reverse) index = cuts.length - 1 - index;
					result.icon = { set: rule.iconSet, index };
					if (rule.showValue === false) result.hideValue = true;
					any = true;
					continue;
				}
				if (!matches(entry, value, row, col)) continue;
				result.style = mergeDxf(result.style, rule.style);
				any = true;
				if ('stopIfTrue' in rule && rule.stopIfTrue) break;
			}
			return any ? result : undefined;
		},
	};
}

/** The colour-scale colour of `value` between threshold positions. */
export function scaleColor(
	value: number,
	cuts: readonly number[],
	colors: readonly string[],
): string | undefined {
	if (!cuts.length || cuts.length !== colors.length) return undefined;
	if (value <= (cuts[0] ?? 0)) return colors[0];
	const last = cuts.length - 1;
	if (value >= (cuts[last] ?? 0)) return colors[last];
	for (let i = 0; i < last; i++) {
		const lo = cuts[i] ?? 0;
		const hi = cuts[i + 1] ?? 0;
		if (value >= lo && value <= hi) {
			const a = colors[i] ?? '#FFFFFF';
			const b = colors[i + 1] ?? a;
			return hi > lo ? mixColors(a, b, (value - lo) / (hi - lo)) : b;
		}
	}
	return colors[last];
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
