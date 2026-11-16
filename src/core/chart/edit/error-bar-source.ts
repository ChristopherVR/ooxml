// Custom error-bar amounts (`c:plus`, `c:minus`, CT_NumDataSource) from what a Format pane's
// source box holds: literal values (`1, 2, 3`, `={1;2;3}`) or a cell range (`Sheet1!$C$2:$C$5`).
// A range keeps its formula and caches the values a resolver reads. The neutral resolver reads only
// cells the chart already caches; a host that holds the workbook (Excel) passes its own.
import type { ChartSpace } from '../model';
import type { ChartDataSource } from '../model-series';
import { formatCellAddress, parseChartFormulaRange } from '../workbook/cellref';
import { chartSeriesEntries } from './draft';

/** Reads the numbers of a `Sheet!Range` formula, or undefined when any cell is unknown. */
export type ChartRangeNumbers = (formula: string) => readonly number[] | undefined;

/** The most points a literal or range error source may name. */
const MAX_POINTS = 10_000;

function cellKeys(formula: string): string[] | undefined {
	const range = parseChartFormulaRange(formula);
	if (!range) return undefined;
	const { start, end } = range;
	const rows = end.row - start.row + 1;
	const cols = end.col - start.col + 1;
	if (rows < 1 || cols < 1 || (rows > 1 && cols > 1) || rows * cols > MAX_POINTS) return undefined;
	const sheet = range.sheetName.toLowerCase();
	return Array.from({ length: rows * cols }, (_, i) =>
		formatCellAddress(
			rows > 1 ? { col: start.col, row: start.row + i } : { col: start.col + i, row: start.row },
		),
	).map((cell) => `${sheet}!${cell}`);
}

function sources(space: ChartSpace): ChartDataSource[] {
	return chartSeriesEntries(space)
		.flatMap(({ series }) => [
			...[
				series.tx?.reference,
				series.categories,
				series.values,
				series.xValues,
				series.yValues,
				series.bubbleSizes,
			],
			...(series.errorBars ?? []).flatMap((bars) => [bars.plus, bars.minus]),
		])
		.filter((source): source is ChartDataSource => !!source?.formula && !!source.cache);
}

/**
 * A resolver over the cells `space` caches: each referenced series name, category, value and
 * error source maps its cached points onto its range. Cells no cache covers are unknown.
 */
export function chartSpaceRangeNumbers(space: ChartSpace): ChartRangeNumbers {
	const cells = new Map<string, string>();
	for (const source of sources(space)) {
		const keys = cellKeys(source.formula!);
		if (!keys) continue;
		source.cache!.points.forEach((point, order) => {
			const key = keys[point.index ?? order];
			if (key && !cells.has(key)) cells.set(key, point.value);
		});
	}
	return (formula) => {
		const keys = cellKeys(formula);
		if (!keys) return undefined;
		const values = keys.map((key) => {
			const text = cells.get(key);
			return text === undefined || text.trim() === '' ? NaN : Number(text);
		});
		return values.every(Number.isFinite) ? values : undefined;
	};
}

/** The text a source box shows for `source`: its formula, else its literal values. */
export function chartErrorBarSourceText(source: ChartDataSource | undefined): string {
	if (!source) return '';
	if (source.formula) return source.formula;
	return (source.cache?.points ?? []).map((point) => point.value).join(', ');
}

const literal = (values: readonly number[]): ChartDataSource => ({
	kind: 'numLit',
	cache: {
		type: 'number',
		formatCode: 'General',
		pointCount: values.length,
		points: values.map((value, index) => ({ index, value: String(value) })),
	},
});

export interface ChartErrorBarSourceOptions {
	/** Points in the series; one literal value is repeated over them, as Office applies it. */
	pointCount: number;
	/** Reads a range's numbers; without it a range is refused. */
	resolveRange?: ChartRangeNumbers;
}

/**
 * The custom error source `text` names: literal nonnegative numbers separated by commas or
 * semicolons (optionally `={...}`), or a single-row or single-column range whose cells resolve to
 * nonnegative numbers. Undefined when the text is neither, so the caller refuses the edit.
 */
export function chartErrorBarSourceFromText(
	text: string,
	options: ChartErrorBarSourceOptions,
): ChartDataSource | undefined {
	const trimmed = text.trim().replace(/^=/u, '').trim();
	if (!trimmed) return undefined;
	const list = /^\{(.*)\}$/u.exec(trimmed)?.[1] ?? trimmed;
	if (/^[\s\d.,;eE+-]+$/u.test(list)) {
		const values = list.split(/[,;]/u).map((part) => (part.trim() === '' ? NaN : Number(part)));
		if (
			!values.length ||
			values.length > MAX_POINTS ||
			!values.every((v) => Number.isFinite(v) && v >= 0)
		)
			return undefined;
		const count = Math.min(Math.max(1, Math.floor(options.pointCount)), MAX_POINTS);
		return literal(values.length === 1 ? Array.from({ length: count }, () => values[0]!) : values);
	}
	if (!cellKeys(trimmed)) return undefined;
	const values = options.resolveRange?.(trimmed);
	if (!values || !values.length || !values.every((v) => Number.isFinite(v) && v >= 0))
		return undefined;
	return {
		kind: 'numRef',
		formula: trimmed,
		cache: literal(values).cache!,
	};
}
