import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../index';
import {
	applyFindReplacements,
	findInSlides,
	replaceInSlides,
} from '../../editor/render/find-replace';
import { createChartTitleDeck } from '../helpers/chart-title-deck';

describe('shared chart title source parts', () => {
	it('reports duplicated title models read-only rather than allowing one to overwrite its sibling', async () => {
		const data = await new PptxHandler().load(await createChartTitleDeck());
		const original = data.slides[0]!.elements[0]!;
		data.slides[0]!.elements.push({ ...structuredClone(original), id: 'duplicate-chart' });
		const matches = findInSlides(data.slides, 'report');
		expect(matches).toHaveLength(2);
		expect(matches.every((match) => match.readOnly)).toBe(true);
		const before = JSON.stringify(data.slides);
		expect(() => replaceInSlides(data.slides, 'report', 'summary')).toThrow('source workbook');
		expect(() =>
			applyFindReplacements(
				data.slides,
				matches.map((match) => ({ ...match, readOnly: false })),
				'summary',
			),
		).toThrow('shared chart');
		expect(JSON.stringify(data.slides)).toBe(before);
	});
	it('retains independent title replacement when chart part paths differ', async () => {
		const data = await new PptxHandler().load(await createChartTitleDeck());
		const original = data.slides[0]!.elements[0]!;
		const duplicate = { ...structuredClone(original), id: 'independent-chart' };
		if (duplicate.type !== 'chart' || !duplicate.chartData) throw new Error('Expected chart');
		duplicate.chartData.chartPartPath = 'ppt/charts/chart2.xml';
		data.slides[0]!.elements.push(duplicate);
		expect(findInSlides(data.slides, 'report').every((match) => !match.readOnly)).toBe(true);
		expect(replaceInSlides(data.slides, 'report', 'summary').replacements).toBe(2);
	});
});
