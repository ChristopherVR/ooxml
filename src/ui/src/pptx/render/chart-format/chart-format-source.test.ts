import type { PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { pptxChartFormatSource } from './chart-format-source';

const bars = (grouping: PptxChartData['grouping']): PptxChartData =>
	({
		chartType: 'bar',
		grouping,
		categories: ['A', 'B'],
		series: [{ name: 'East', values: [1, 2] }],
	}) as PptxChartData;

describe('pptxChartFormatSource label options', () => {
	it('lists the positions the chart model accepts', () => {
		const positions = (data: PptxChartData) =>
			pptxChartFormatSource(data).labelOptions?.(0)?.positions;
		expect(positions(bars('clustered'))).toContain('outEnd');
		// Office has no Outside End for stacked bars, and the core refuses it.
		expect(positions(bars('stacked'))).toEqual(['ctr', 'inEnd', 'inBase']);
	});
});
