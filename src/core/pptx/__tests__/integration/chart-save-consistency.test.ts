import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../index';
import type { ChartPptxElement } from '../../index';
import { describeChartSaveConsistency, wholeColumnFixture } from './chart-save-consistency';

describeChartSaveConsistency('whole-column', 'history');

describe('category labels of a chart that names none', () => {
	it('are not read from the first column of its workbook', async () => {
		// Excel labels such a chart 1, 2, 3; here the first column holds a series of its own.
		const { bytes } = await wholeColumnFixture(false);
		const handler = new PptxHandler();
		const open = () =>
			handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
			);
		const { slides } = await open();
		const data = slides
			.flatMap((slide) => slide.elements)
			.find(
				(element): element is ChartPptxElement =>
					element.type === 'chart' && element.chartData?.chartPartPath === 'ppt/charts/whole.xml',
			)!.chartData!;
		expect(data.categories).toEqual([]);
		expect(data.series[0]!.name).toBe('Labels');
		handler.dispose();
	});
});
