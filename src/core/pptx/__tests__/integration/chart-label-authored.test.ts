import { expect, it } from 'vitest';
import { PresentationBuilder, PptxHandler } from '../../index';
import { findInSlides, replaceInSlides } from '../../editor/render/find-replace';
import { createChartLabelDeck } from '../helpers/chart-label-deck';

it('rejects label replacements after a structural edit invalidates the loaded source grid', async () => {
	const handler = new PptxHandler();
	const data = await handler.load(await createChartLabelDeck());
	const element = data.slides[0]!.elements.find((item) => item.type === 'chart')!;
	if (element.type !== 'chart') throw new Error('Expected chart');
	expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]?.readOnly).toBe(false);
	element.chartData!.categories.push('Q3');
	element.chartData!.series[0]!.values.push(30);
	expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]?.readOnly).toBe(true);
	expect(() =>
		replaceInSlides(data.slides, 'Dataset', 'Renamed', { includeChartData: true }),
	).toThrow();
	handler.dispose();
});

it('replaces newly authored literal labels before their first save and preserves them on reopen', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [
		createSlide('Blank')
			.addChart(
				'bar',
				{
					categories: ['Q1', 'Q2'],
					series: [{ name: 'Dataset', values: [10, 20] }],
				},
				{ x: 80, y: 60, width: 700, height: 400 },
			)
			.build(),
	];
	expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]?.readOnly).toBe(false);
	let result = replaceInSlides(data.slides, 'Dataset', 'Renamed', { includeChartData: true });
	result = replaceInSlides([...result.slides], 'Q1', 'First quarter', { includeChartData: true });
	const bytes = await handler.save([...result.slides]);
	const reopenedHandler = new PptxHandler();
	const reopened = await reopenedHandler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	expect(findInSlides(reopened.slides, 'Renamed', { includeChartData: true })[0]?.readOnly).toBe(
		false,
	);
	expect(findInSlides(reopened.slides, 'First quarter', { includeChartData: true })).toHaveLength(
		1,
	);
	expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })).toHaveLength(1);
	handler.dispose();
	reopenedHandler.dispose();
});

it('does not infer editable literal sources for authored external or extended charts', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [
		createSlide('Blank')
			.addChart(
				'bar',
				{
					categories: ['Q1'],
					series: [{ name: 'Dataset', values: [10] }],
				},
				{ x: 80, y: 60, width: 700, height: 400 },
			)
			.build(),
	];
	const element = data.slides[0]!.elements[0]!;
	if (element.type !== 'chart') throw new Error('Expected chart');
	element.chartData!.externalData = { relationshipId: 'rId1', isExternal: true };
	expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]?.readOnly).toBe(true);
	delete element.chartData!.externalData;
	element.chartData!.chartType = 'waterfall';
	expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]?.readOnly).toBe(true);
	handler.dispose();
});
