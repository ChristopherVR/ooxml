/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import { PptxHandler, type ChartPptxElement } from 'pptx-viewer-core';

import { loadDeck, slideStage } from './support/deck';

test('renders rectangular chart gradients as stretched image patterns', async ({ page }, info) => {
	const { handler, data, createSlide } = await PptxHandler.createBlank();
	const slide = createSlide('Blank')
		.addChart('bar', {
			categories: ['A', 'B'],
			series: [{ name: 'Series', values: [1, 2] }],
		})
		.build();
	const chart = slide.elements.find((element) => element.type === 'chart') as ChartPptxElement;
	const gradient = {
		type: 'radial' as const,
		path: 'rect',
		stops: [
			{ color: '#ff0000', position: 0 },
			{ color: '#0000ff', position: 100 },
		],
	};
	chart.chartData!.style = { ...chart.chartData!.style, chartAreaGradient: gradient };
	chart.chartData!.series[0]!.gradientFill = gradient;
	data.slides.push(slide);
	const sourcePath = info.outputPath('chart-rectpath.pptx');
	await writeFile(sourcePath, await handler.save(data.slides));
	await loadDeck(page, sourcePath);
	const patterns = slideStage(page).locator('pattern[patternContentUnits="objectBoundingBox"]');
	await expect(patterns).toHaveCount(2);
	for (const pattern of await patterns.all()) {
		await expect(pattern).toHaveAttribute('patternUnits', 'objectBoundingBox');
		await expect(pattern).toHaveAttribute('width', '1');
		await expect(pattern).toHaveAttribute('height', '1');
		const image = pattern.locator('image');
		await expect(image).toHaveAttribute('preserveAspectRatio', 'none');
		await expect(image).toHaveAttribute('href', /^data:image\/svg\+xml,/u);
	}
});
