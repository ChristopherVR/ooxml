import { describe, expect, it } from 'vitest';
import { aggregateByCategory, computeHistogramBins, scottBinCount } from './histogram-binning.js';

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
