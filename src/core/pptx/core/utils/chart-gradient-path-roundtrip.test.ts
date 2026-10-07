import { expect, it } from 'vitest';
import { PptxHandler } from '../PptxHandler';
import type { ChartPptxElement, PptxSlide } from '../types';

const chartOf = (slides: PptxSlide[]) =>
	slides[0]!.elements.find((element) => element.type === 'chart') as ChartPptxElement;

it('retains rectangular chart paths across saves and applies a change to circle', async () => {
	const blank = await PptxHandler.createBlank();
	const slide = blank
		.createSlide('Blank')
		.addChart('bar', {
			categories: ['A', 'B'],
			series: [{ name: 'Series', values: [1, 2] }],
		})
		.build();
	let slides = [slide];
	let handler = blank.handler;
	const gradient = {
		type: 'radial' as const,
		path: 'rect',
		stops: [
			{ color: '#ff0000', position: 0 },
			{ color: '#0000ff', position: 100 },
		],
	};
	chartOf(slides).chartData!.style = { chartAreaGradient: gradient };
	chartOf(slides).chartData!.series[0]!.gradientFill = { ...gradient };
	for (let cycle = 0; cycle < 2; cycle++) {
		const bytes = await handler.save(slides);
		handler = new PptxHandler();
		slides = (
			await handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
			)
		).slides;
		expect(chartOf(slides).chartData!.style!.chartAreaGradient!.path).toBe('rect');
		expect(chartOf(slides).chartData!.series[0]!.gradientFill!.path).toBe('rect');
	}
	chartOf(slides).chartData!.style!.chartAreaGradient!.path = 'circle';
	const bytes = await handler.save(slides);
	const loaded = await new PptxHandler().load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	expect(chartOf(loaded.slides).chartData!.style!.chartAreaGradient!.path).toBeUndefined();
	expect(chartOf(loaded.slides).chartData!.series[0]!.gradientFill!.path).toBe('rect');
});
