import { describe, expect, it } from 'vitest';

import { CHART_AUTO_TITLE, impliesAutoTitle, resolveAutoTitleText } from './auto-title';

describe('auto title', () => {
	it('titles a single named series chart with that name', () => {
		const input = { chartType: 'pie', seriesNames: ['Sales'] };
		expect(impliesAutoTitle(input)).toBe(true);
		expect(resolveAutoTitleText({ ...input, hasTitle: true })).toBe('Sales');
	});

	it('uses the placeholder for several series and for chartex families', () => {
		expect(
			resolveAutoTitleText({ chartType: 'bar', seriesNames: ['A', 'B'], hasTitle: true }),
		).toBe(CHART_AUTO_TITLE);
		expect(resolveAutoTitleText({ chartType: 'funnel', seriesNames: ['A'], hasTitle: true })).toBe(
			CHART_AUTO_TITLE,
		);
	});

	it('never auto-titles when the title was toggled or deleted', () => {
		expect(impliesAutoTitle({ chartType: 'bar', seriesNames: ['A'], hasTitle: false })).toBe(false);
		expect(impliesAutoTitle({ chartType: 'bar', seriesNames: ['A'], autoTitleDeleted: true })).toBe(
			false,
		);
		expect(impliesAutoTitle({ chartType: 'waterfall', seriesNames: ['A'] })).toBe(false);
	});

	it('prefers an explicit title and shows nothing when hidden', () => {
		expect(
			resolveAutoTitleText({ chartType: 'bar', seriesNames: ['A'], hasTitle: true, title: 'T' }),
		).toBe('T');
		expect(
			resolveAutoTitleText({ chartType: 'bar', seriesNames: ['A'], title: 'T' }),
		).toBeUndefined();
	});
});
