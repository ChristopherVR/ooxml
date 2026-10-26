import { describe, expect, it } from 'vitest';

import { distributeRunsText, realignOrCollapseRuns } from './realign-runs';

describe('distributeRunsText', () => {
	it('returns undefined for an empty run list', () => {
		expect(distributeRunsText([], 'anything')).toBeUndefined();
	});

	it('keeps a single unchanged run and appends trailing text to it', () => {
		expect(distributeRunsText(['Revenue'], 'Revenue Growth')).toStrictEqual(['Revenue Growth']);
	});

	it('distributes an append across two unchanged runs onto the last one', () => {
		expect(distributeRunsText(['Revenue', ' Growth'], 'Revenue Growth 2024')).toStrictEqual([
			'Revenue',
			' Growth 2024',
		]);
	});

	it('confines a middle-run rewrite to that run alone', () => {
		expect(
			distributeRunsText(['Revenue ', 'Growth', ' Report'], 'Revenue Increase Report'),
		).toStrictEqual(['Revenue ', 'Increase', ' Report']);
	});

	it('returns undefined for a two-run title whose LAST run is edited (no later run to anchor on)', () => {
		// With only two runs, a change to the last run's own text has no
		// subsequent run boundary to realign against, so this degrades to the
		// same fallback an unrelated rewrite gets; the three-run case (a
		// changed run sandwiched between two unchanged ones) is what
		// `distributes ... onto that run alone` above covers.
		expect(distributeRunsText(['Q4', ' Sales'], 'Q4 Total Sales')).toBeUndefined();
	});

	it('returns undefined when no later run boundary survives (an unrelated rewrite)', () => {
		expect(
			distributeRunsText(['Revenue', ' Growth'], 'Completely Different Title'),
		).toBeUndefined();
	});

	it('returns undefined when only the first of three runs still resolves', () => {
		// 'Growth' and 'Report' both vanish: no anchor to realign the tail onto.
		expect(
			distributeRunsText(['Revenue ', 'Growth', ' Report'], 'Revenue Something Else'),
		).toBeUndefined();
	});

	it('clears an intermediate run absorbed entirely by an edit spanning past it', () => {
		// Run 'B' vanishes and 'C' is nowhere to be found either, but 'D'
		// resurfaces: the whole 'xyz' gap lands on run 1 (the first
		// non-matching run), run 2 ('C') is cleared since it was skipped over,
		// and run 3 ('D') resumes as an unchanged anchor.
		expect(distributeRunsText(['A', 'B', 'C', 'D'], 'AxyzD')).toStrictEqual(['A', 'xyz', '', 'D']);
	});
});

describe('realignOrCollapseRuns', () => {
	it('realigns onto the existing runs when the new text still contains them', () => {
		const runs = realignOrCollapseRuns(
			[{ text: 'Revenue', bold: true }, { text: ' Growth' }],
			'Revenue Growth 2024',
		);
		expect(runs).toStrictEqual([{ text: 'Revenue', bold: true }, { text: ' Growth 2024' }]);
	});

	it("collapses to a single run in the first run's formatting on an unrelated full rewrite", () => {
		const runs = realignOrCollapseRuns(
			[
				{ text: 'Revenue', bold: true, color: '#FF0000' },
				{ text: ' Growth', italic: true },
			],
			'Completely Different Title',
		);
		expect(runs).toStrictEqual([
			{ text: 'Completely Different Title', bold: true, color: '#FF0000' },
		]);
	});

	it("collapses three stale runs to one, keeping only the first run's formatting", () => {
		const runs = realignOrCollapseRuns(
			[{ text: 'Q1 ', italic: true }, { text: 'Revenue', bold: true }, { text: ' Report' }],
			'Annual Summary',
		);
		expect(runs).toStrictEqual([{ text: 'Annual Summary', italic: true }]);
	});
});
