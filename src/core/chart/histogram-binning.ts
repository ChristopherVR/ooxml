/**
 * Raw-observation binning for the histogram/Pareto ChartEx series: numeric
 * range binning (`cx:layoutPr/cx:binning`) and category aggregation
 * (`cx:layoutPr/cx:aggregation`). Format-neutral: the bin labels are built with
 * a caller-supplied number formatter, so the viewer keeps its own axis
 * formatting.
 */
import { groupRowsByCategory } from './box-stats';
import { extentOf } from './extent';

/** The `cx:layoutPr/cx:binning` properties that drive numeric binning. */
export interface HistogramOptions {
	binSize?: number | undefined;
	binCount?: number | undefined;
	intervalClosed?: 'l' | 'r' | undefined;
	underflow?: number | 'auto' | undefined;
	overflow?: number | 'auto' | undefined;
}

/** Formats a bin edge for a label. */
export type HistogramValueFormatter = (value: number) => string;

export interface HistogramBin {
	value: number;
	label: string;
	sourceIndices: number[];
}

function binLabel(
	lower: number,
	upper: number,
	closed: 'l' | 'r',
	format: HistogramValueFormatter,
): string {
	return closed === 'r'
		? `(${format(lower)}, ${format(upper)}]`
		: `[${format(lower)}, ${format(upper)})`;
}

/**
 * Excel/PowerPoint's automatic histogram bin count when the author specifies
 * neither `binSize` nor `binCount` (Scott's normal reference rule: bin width
 * = 3.49 * sample standard deviation * n^(-1/3), then the count needed to
 * span the data at that width). COM-verified against charts-com.pptx slide
 * 30 / chartEx5.xml: 76 raw values spanning 1-24 (sample sd ~4.898) gave
 * PowerPoint's own 6 bins of width 4 each; Scott's rule on the same data
 * gives a raw width of ~4.035, which needs exactly 6 bins to span the range,
 * matching PowerPoint bin-for-bin. The previous `ceil(sqrt(n))` fallback
 * gave 9 bins on the same data, visibly narrower than PowerPoint's own.
 */
export function scottBinCount(values: ReadonlyArray<number>, min: number, max: number): number {
	if (values.length < 2 || max <= min) {
		return 1;
	}
	const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
	const variance =
		values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
	const sd = Math.sqrt(variance);
	if (!(sd > 0)) {
		return 1;
	}
	const width = 3.49 * sd * values.length ** (-1 / 3);
	return Math.max(1, Math.ceil((max - min) / width));
}

/**
 * Most bins one histogram draws. `cx:binCount` and `cx:binSize` come from the file and the bins
 * are built one by one, so an unbounded request exhausts memory; a chart cannot show more bars
 * than this anyway.
 */
export const MAX_HISTOGRAM_BINS = 1000;

interface Observation {
	value: number;
	sourceIndex: number;
}

/** Finite observations, split into the underflow, overflow and regular bins (linear time). */
function partitionObservations(
	values: ReadonlyArray<number>,
	closed: 'l' | 'r',
	underflow: number | undefined,
	overflow: number | undefined,
) {
	const under: Observation[] = [];
	const over: Observation[] = [];
	const regular: Observation[] = [];
	values.forEach((value, sourceIndex) => {
		if (!Number.isFinite(value)) return;
		const item = { value, sourceIndex };
		const isUnder =
			underflow !== undefined && (closed === 'r' ? value <= underflow : value < underflow);
		const isOver =
			overflow !== undefined && (closed === 'l' ? value >= overflow : value > overflow);
		if (isUnder) under.push(item);
		if (isOver) over.push(item);
		if (!isUnder && !isOver) regular.push(item);
	});
	return { under, over, regular };
}

/** Where the regular bins fall, with `limited` set when the request passed the bin limit. */
function binGrid(values: ReadonlyArray<number>, options: HistogramOptions, closed: 'l' | 'r') {
	const { min, max } = extentOf(values);
	const size =
		options.binSize && options.binSize > 0 && Number.isFinite(options.binSize)
			? options.binSize
			: undefined;
	if (size !== undefined) {
		const start = Math.floor(min / size) * size;
		const wanted =
			closed === 'l' ? Math.floor((max - start) / size) + 1 : Math.ceil((max - start) / size);
		if (Number.isFinite(start) && wanted <= MAX_HISTOGRAM_BINS)
			return { start, width: size, count: Math.max(wanted, 1), limited: false };
	}
	// A bin size too fine for the limit spreads the most bins allowed over the data instead.
	const wanted =
		size !== undefined
			? Number.POSITIVE_INFINITY
			: options.binCount && options.binCount > 0
				? options.binCount
				: scottBinCount(values, min, max);
	const count = Math.max(Math.min(Math.floor(wanted) || 1, MAX_HISTOGRAM_BINS), 1);
	return {
		start: min,
		width: Math.max((max - min) / count, 1),
		count,
		limited: wanted > MAX_HISTOGRAM_BINS,
	};
}

const numericLimit = (limit: number | 'auto' | undefined) =>
	typeof limit === 'number' ? limit : undefined;

/** True when {@link computeHistogramBins} draws fewer bins than `options` ask for. */
export function histogramBinsLimited(
	values: ReadonlyArray<number>,
	options: HistogramOptions,
): boolean {
	const closed = options.intervalClosed ?? 'l';
	const { regular } = partitionObservations(
		values,
		closed,
		numericLimit(options.underflow),
		numericLimit(options.overflow),
	);
	if (regular.length === 0) return false;
	return binGrid(
		regular.map((item) => item.value),
		options,
		closed,
	).limited;
}

/**
 * Bin raw observations according to ChartEx binning properties, drawing at most
 * {@link MAX_HISTOGRAM_BINS} regular bins.
 */
export function computeHistogramBins(
	values: ReadonlyArray<number>,
	options: HistogramOptions,
	format: HistogramValueFormatter = String,
): HistogramBin[] {
	const closed = options.intervalClosed ?? 'l';
	const underflow = numericLimit(options.underflow);
	const overflow = numericLimit(options.overflow);
	const { under, over, regular } = partitionObservations(values, closed, underflow, overflow);
	if (under.length + over.length + regular.length === 0) {
		return [];
	}
	const result: HistogramBin[] = [];
	if (underflow !== undefined) {
		result.push({
			value: under.length,
			label: `${closed === 'r' ? '≤' : '<'} ${format(underflow)}`,
			sourceIndices: under.map((item) => item.sourceIndex),
		});
	}
	if (regular.length > 0) {
		const { start, width, count } = binGrid(
			regular.map((item) => item.value),
			options,
			closed,
		);
		const bins = Array.from({ length: count }, (_, index) => ({
			value: 0,
			label: binLabel(start + index * width, start + (index + 1) * width, closed, format),
			sourceIndices: [] as number[],
		}));
		for (const item of regular) {
			const rawIndex =
				closed === 'r'
					? Math.ceil((item.value - start) / width) - 1
					: Math.floor((item.value - start) / width);
			const index = Math.max(0, Math.min(rawIndex, bins.length - 1));
			const bin = bins[index];
			if (!bin) continue;
			bin.value += 1;
			bin.sourceIndices.push(item.sourceIndex);
		}
		result.push(...bins);
	}
	if (overflow !== undefined) {
		result.push({
			value: over.length,
			label: `${closed === 'l' ? '≥' : '>'} ${format(overflow)}`,
			sourceIndices: over.map((item) => item.sourceIndex),
		});
	}
	return result;
}

/**
 * Aggregate raw observations by their (repeated) category label instead of by
 * numeric value range: `cx:layoutPr/cx:aggregation` (COM-verified against
 * charts-com.pptx slide 31 / chartEx6.xml) means the source rows are
 * CATEGORICAL, one row per occurrence, with the value column carrying a
 * constant `1` so summing it counts occurrences per category. Summing
 * (rather than literally counting rows) also does the right thing for a
 * dataset that authors a real numeric measure per row instead of a bare `1`.
 */
export function aggregateByCategory(
	values: ReadonlyArray<number>,
	rawCategories: ReadonlyArray<string>,
): HistogramBin[] {
	const { uniqueCategories, rowIndexesByCategory } = groupRowsByCategory(rawCategories);
	return uniqueCategories.map((label) => {
		const rows = rowIndexesByCategory.get(label) ?? [];
		return {
			value: rows.reduce((sum, row) => sum + (values[row] ?? 0), 0),
			label,
			sourceIndices: rows,
		};
	});
}
