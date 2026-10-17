import { ensureBarFacePicturePixelSampled, releasePptxRenderCaches } from 'ooxml-ui/pptx';
import { disposePptxRenderStage } from 'ooxml-ui/pptx/dom';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { PptxViewer } from './PptxViewer';

const slide: PptxSlide = {
	id: 'slide',
	rId: 'rId1',
	slideNumber: 1,
	elements: [
		{
			type: 'chart',
			id: 'chart',
			x: 0,
			y: 0,
			width: 400,
			height: 300,
			chartData: { chartType: 'bar', categories: ['A'], series: [{ name: 'One', values: [1] }] },
		},
	],
};
let viewer: PptxViewer | undefined;
const parking = document.createElement('div');
let sampleIndex = 0;
async function sample(): Promise<void> {
	ensureBarFacePicturePixelSampled(`blob:vanilla-${sampleIndex++}`, async () => '#123456');
	for (let i = 0; i < 6; i++) await Promise.resolve();
}
afterEach(async () => {
	viewer?.destroy();
	viewer = undefined;
	disposePptxRenderStage(parking);
	parking.replaceChildren();
	document.body.replaceChildren();
	releasePptxRenderCaches();
	await new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});
});
function mount(): HTMLElement {
	const container = document.createElement('div');
	document.body.append(container, parking);
	viewer = new PptxViewer(container);
	viewer.store.set({ slides: [slide] });
	return container;
}

describe('vanilla rendered stage ownership', () => {
	it('retires main and thumbnail chart subscriptions when stages are replaced', async () => {
		const container = mount();
		const charts = [...container.querySelectorAll<HTMLElement>('.pptxv-chart')];
		expect(charts.length).toBeGreaterThanOrEqual(2);
		const svgs = charts.map((chart) => chart.querySelector('svg'));
		viewer!.store.set({ canvasSize: { width: 960, height: 540 } });
		for (const chart of charts) {
			expect(chart.isConnected).toBe(false);
			parking.append(chart);
		}
		await sample();
		for (const [index, chart] of charts.entries())
			expect(chart.querySelector('svg')).toBe(svgs[index]);
	});

	it('retires charts when the complete viewer chrome is destroyed', async () => {
		const container = mount();
		const charts = [...container.querySelectorAll<HTMLElement>('.pptxv-chart')];
		const svgs = charts.map((chart) => chart.querySelector('svg'));
		viewer!.destroy();
		for (const chart of charts) parking.append(chart);
		await sample();
		for (const [index, chart] of charts.entries())
			expect(chart.querySelector('svg')).toBe(svgs[index]);
	});
});
