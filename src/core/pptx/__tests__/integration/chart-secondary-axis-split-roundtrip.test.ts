// Plot Series On > Secondary Axis for the last series of a stacked column chart PowerPoint wrote:
// the save must split it into its own bar group on a new axis pair (PowerPoint 16 otherwise reads
// it on the primary axis), keep the other two stacked, and move it back on a later save.
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../../../chart/index';
import { withChartSeriesOnSecondaryAxis } from '../../core/builders/sdk/chart-series-axis-operations';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement } from '../../core/types';

const FIXTURE = new URL('../fixtures/e2e/chart-data-fidelity.pptx', import.meta.url);

async function columnChart(bytes: Uint8Array) {
	const handler = new PptxHandler();
	const deck = await handler.load(bytes);
	const slide = deck.slides.findIndex((item) =>
		item.elements.some((e) => e.type === 'chart' && e.chartData?.chartType === 'bar'),
	);
	const element = deck.slides[slide]!.elements.find((e) => e.type === 'chart') as ChartPptxElement;
	return { handler, deck, element };
}

async function groups(bytes: Uint8Array, part: string) {
	const zip = await JSZip.loadAsync(bytes);
	const space = parseChartSpace(await zip.file(part)!.async('string')).chartSpace;
	return space.plotArea.groups.map((group) => ({
		kind: group.kind,
		grouping: group.grouping,
		series: group.series.length,
		valueAxis: space.plotArea.axes.find((axis) => axis.id === group.axisIds[1])?.position,
	}));
}

describe('PowerPoint secondary axis split', () => {
	it('saves a stacked end series on its own group and secondary pair, and back', async () => {
		const first = await columnChart(new Uint8Array(readFileSync(FIXTURE)));
		first.element.chartData!.categories = ['Plan', 'Design', 'Build', 'Review', 'Ship'];
		first.element.chartData!.series = [
			{ name: 'Series 1', values: [2, 4, 6, 8, 10] },
			{ name: 'Series 2', values: [3, 6, 9, 12, 15] },
			{ name: 'Series 3', values: [1, 5, 7, 11, 13] },
		];
		first.element.chartData = withChartSeriesOnSecondaryAxis(
			{ ...first.element.chartData!, grouping: 'stacked' },
			2,
			true,
		);
		const part = first.element.chartData.chartPartPath!;
		const saved = await first.handler.save(first.deck.slides);
		expect(await groups(saved, part)).toEqual([
			{ kind: 'bar', grouping: 'stacked', series: 2, valueAxis: 'l' },
			{ kind: 'bar', grouping: 'stacked', series: 1, valueAxis: 'r' },
		]);

		const second = await columnChart(saved);
		const moved = second.element.chartData!;
		expect(moved.series.map((s) => s.values)).toEqual([
			[2, 4, 6, 8, 10],
			[3, 6, 9, 12, 15],
			[1, 5, 7, 11, 13],
		]);
		second.element.chartData = withChartSeriesOnSecondaryAxis(moved, 2, false);
		const back = await second.handler.save(second.deck.slides);
		expect(await groups(back, part)).toEqual([
			{
				kind: 'bar',
				grouping: 'stacked',
				series: 3,
				valueAxis: 'l',
			},
		]);
	});
});
