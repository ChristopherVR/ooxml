// Trendlines and error bars set through PowerPoint's Format pane operations survive a save and
// reopen of a PowerPoint-authored deck.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import { chartErrorBarSourceFromText } from '../../../../chart/index';
import { PptxHandler } from '../../PptxHandler';
import type { ChartPptxElement } from '../../types';
import { withChartSeriesErrorBars } from './chart-error-bar-operations';
import { chartSeriesTrendlinesView, withChartSeriesTrendline } from './chart-trendline-operations';

const FIXTURE = new URL(
	'../../../__tests__/fixtures/e2e/chart-data-fidelity.pptx',
	import.meta.url,
);

const columnOf = (slides: Awaited<ReturnType<PptxHandler['load']>>['slides']) =>
	slides
		.flatMap((slide) => slide.elements)
		.find(
			(item) => item.type === 'chart' && item.chartData?.chartType === 'bar',
		) as ChartPptxElement;

describe('PowerPoint chart analysis saves', () => {
	it('reopens several trendlines and custom error values', async () => {
		const handler = new PptxHandler();
		const deck = await handler.load(new Uint8Array(readFileSync(FIXTURE)));
		const element = columnOf(deck.slides);
		const points = element.chartData!.series[0]!.values.length;
		let data = withChartSeriesTrendline(element.chartData!, 0, 0, { type: 'linear' });
		data = withChartSeriesTrendline(data, 0, 1, { type: 'poly', order: 3, displayRSquared: true });
		data = withChartSeriesErrorBars(data, 0, {
			valueType: 'cust',
			plus: chartErrorBarSourceFromText('2', { pointCount: points })!,
			minus: chartErrorBarSourceFromText('1', { pointCount: points })!,
		});
		element.chartData = data;
		const reopened = columnOf(
			(await new PptxHandler().load(await handler.save(deck.slides))).slides,
		);
		const back = reopened.chartData!;
		expect(chartSeriesTrendlinesView(back, 0)).toEqual([
			{ trendlineType: 'linear' },
			{ trendlineType: 'poly', order: 3, displayRSquared: true },
		]);
		expect(back.series[0]!.errBars?.find((bars) => bars.direction === 'y')).toMatchObject({
			valType: 'cust',
			customPlus: Array.from({ length: points }, () => 2),
			customMinus: Array.from({ length: points }, () => 1),
		});
	});
});
