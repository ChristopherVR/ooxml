import { describe, expect, it } from 'vitest';
import {
	MAX_HISTOGRAM_BINS,
	aggregateByCategory,
	computeHistogramBins,
	histogramBinsLimited,
	scottBinCount,
} from './histogram-binning';

describe('histogram bin limit', () => {
	const values = [1, 2, 3];

	it.each([
		['a huge bin count', { binCount: 200_000_000 }],
		['an infinite bin count', { binCount: Number.POSITIVE_INFINITY }],
		['a tiny bin size', { binSize: 1e-9 }],
		['a subnormal bin size', { binSize: 5e-324 }],
	])('draws at most the limit for %s and keeps every value', (_, options) => {
		const bins = computeHistogramBins(values, options);
		expect(bins).toHaveLength(MAX_HISTOGRAM_BINS);
		expect(bins.reduce((sum, bin) => sum + bin.value, 0)).toBe(values.length);
		expect(histogramBinsLimited(values, options)).toBe(true);
		for (const bin of [bins[0], bins.at(-1)]) expect(bin?.label).not.toMatch(/Infinity|NaN/);
	});

	it('spreads a limited bin size over the whole range', () => {
		const bins = computeHistogramBins([0, 5000], { binSize: 0.001 });
		expect(bins).toHaveLength(MAX_HISTOGRAM_BINS);
		expect(bins[0]?.label).toBe('[0, 5)');
		expect(bins.at(-1)?.label).toBe('[4995, 5000)');
		expect(bins.at(-1)?.value).toBe(1);
	});

	it('leaves requests within the limit alone', () => {
		expect(computeHistogramBins([0, 999], { binSize: 1 })).toHaveLength(MAX_HISTOGRAM_BINS);
		expect(histogramBinsLimited([0, 999], { binSize: 1 })).toBe(false);
		expect(computeHistogramBins(values, { binCount: MAX_HISTOGRAM_BINS })).toHaveLength(
			MAX_HISTOGRAM_BINS,
		);
		expect(histogramBinsLimited(values, { binCount: MAX_HISTOGRAM_BINS })).toBe(false);
		expect(histogramBinsLimited(values, {})).toBe(false);
	});

	it('ignores values that fall in the overflow and underflow bins', () => {
		expect(histogramBinsLimited([1], { binCount: 1e9, underflow: 5 })).toBe(false);
	});

	it('splits many observations in linear time', () => {
		const many = Array.from({ length: 200_000 }, (_, index) => index);
		const started = performance.now();
		const bins = computeHistogramBins(many, { binCount: 4, underflow: 100_000, overflow: 100_001 });
		expect(performance.now() - started).toBeLessThan(2000);
		expect(bins[0]?.value).toBe(100_000);
		expect(bins.at(-1)?.value).toBe(99_999);
	});
});

describe('scottBinCount', () => {
	it('is one bin for degenerate samples', () => {
		expect(scottBinCount([], 0, 0)).toBe(1);
		expect(scottBinCount([3, 3, 3], 3, 3)).toBe(1);
	});

	it('spans the range at the normal reference width', () => {
		const values = Array.from({ length: 76 }, (_, i) => 1 + ((i * 7) % 24));
		expect(scottBinCount(values, 1, 24)).toBeGreaterThanOrEqual(2);
	});
});

describe('computeHistogramBins', () => {
	it('bins by an explicit size with left-closed intervals', () => {
		const bins = computeHistogramBins([1, 2, 3, 4, 5, 6], { binSize: 2 });
		expect(bins.map((bin) => bin.label)).toEqual(['[0, 2)', '[2, 4)', '[4, 6)', '[6, 8)']);
		expect(bins.map((bin) => bin.value)).toEqual([1, 2, 2, 1]);
	});

	it('keeps source indices and skips non-finite values', () => {
		const bins = computeHistogramBins([1, Number.NaN, 3], { binSize: 2 });
		expect(bins.flatMap((bin) => bin.sourceIndices).sort()).toEqual([0, 2]);
	});

	it('bins more values than a call can pass as arguments', () => {
		const values = Array.from({ length: 200_000 }, (_, i) => 1 + (i % 9));
		values[150_000] = 0;
		values[60_000] = 10;
		const bins = computeHistogramBins(values, { binSize: 5 });
		expect(bins.map((bin) => bin.label)).toEqual(['[0, 5)', '[5, 10)', '[10, 15)']);
		expect(bins.reduce((sum, bin) => sum + bin.value, 0)).toBe(200_000);
	});

	it('closes intervals on the right when asked', () => {
		const bins = computeHistogramBins([2, 4], { binSize: 2, intervalClosed: 'r' });
		expect(bins.map((bin) => bin.label)).toEqual(['(2, 4]']);
		expect(bins.map((bin) => bin.value)).toEqual([2]);
	});

	it('adds underflow and overflow bins and uses the injected formatter', () => {
		const bins = computeHistogramBins(
			[1, 5, 9],
			{ binCount: 1, underflow: 2, overflow: 8 },
			(value) => `${value}u`,
		);
		expect(bins.map((bin) => bin.label)).toEqual(['< 2u', '[5u, 6u)', '≥ 8u']);
		expect(bins.map((bin) => bin.value)).toEqual([1, 1, 1]);
	});

	it('returns nothing without finite values', () => {
		expect(computeHistogramBins([Number.NaN], {})).toEqual([]);
	});
});

describe('aggregateByCategory', () => {
	it('sums the value column per repeated category in first-seen order', () => {
		const bins = aggregateByCategory([1, 1, 1, 2], ['b', 'a', 'b', 'a']);
		expect(bins).toEqual([
			{ value: 2, label: 'b', sourceIndices: [0, 2] },
			{ value: 3, label: 'a', sourceIndices: [1, 3] },
		]);
	});
});
