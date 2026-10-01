import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	PresentationBuilder,
	setChartDataPointExplosion,
	setChartDataPointFill,
} from '../../core/builders/sdk';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement } from '../../core/types/elements';

type Loaded = Awaited<ReturnType<PptxHandler['load']>>;

function chartOf(data: Loaded): ChartPptxElement {
	return data.slides[0].elements.find((element) => element.type === 'chart') as ChartPptxElement;
}

async function chartXml(bytes: Uint8Array): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file('ppt/charts/chart1.xml')!.async('string');
}

async function editAndSave(bytes: Uint8Array, edit: (chart: ChartPptxElement) => void) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes.buffer as ArrayBuffer);
	edit(chartOf(data));
	data.slides[0].isDirty = true;
	return handler.save(data.slides);
}

describe('clearing the last c:dPt override of a loaded chart', () => {
	it('removes the authored c:dPt from the saved file', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(
			createSlide('Blank')
				.addChart(
					'pie',
					{ categories: ['A', 'B', 'C'], series: [{ name: 'S', values: [1, 2, 3] }] },
					{ x: 20, y: 20, width: 400, height: 300 },
				)
				.build(),
		);
		const saved = await handler.save(data.slides);

		const withPoint = await editAndSave(saved, (chart) => {
			setChartDataPointFill(chart, 0, 1, '#FF0000');
			setChartDataPointExplosion(chart, 0, 2, 15);
		});
		expect((await chartXml(withPoint)).match(/<c:dPt>/gu)).toHaveLength(2);

		const cleared = await editAndSave(withPoint, (chart) => {
			setChartDataPointFill(chart, 0, 1, null);
			setChartDataPointExplosion(chart, 0, 2, null);
		});
		await expect(chartXml(cleared)).resolves.not.toContain('<c:dPt>');
	});
});
