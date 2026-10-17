// @vitest-environment happy-dom
/**
 * Regression test for the async first-pixel colour sample cache
 * (`chart-bar3d-face-picture-sample.ts`): an untargeted bar3D extrusion face
 * whose fill is picture-only paints the resolved point/series colour until
 * the picture's first pixel decodes, then repaints once it lands - and the
 * repaint subscription self-unsubscribes once the chart's container has left
 * the document (vanilla has no unmount hook to drive cleanup from).
 */
import type { PptxChartData, PptxElement } from 'ooxml-core/pptx';
import {
	ensureBarFacePicturePixelSampled,
	getCachedBarFacePicturePixelColor,
	resetBarFacePicturePixelCacheForTests,
	resolveUntargetedBarFaceFill,
} from '../../index';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n/translator';
import { createElementRendererRegistry } from '../registry';
import type { ElementRenderContext } from '../types';
import { renderChartElement } from './chart';

const IMAGE_URL = 'data:image/png;base64,vanilla-sample-test-image';
const SAMPLED_COLOR = '#0e0e0e';

beforeEach(() => {
	resetBarFacePicturePixelCacheForTests();
});
afterEach(() => {
	resetBarFacePicturePixelCacheForTests();
});

function buildContext(): ElementRenderContext {
	const registry = createElementRendererRegistry();
	const context: ElementRenderContext = {
		document,
		slide: { id: 'slide-1', rId: 'rId1', slideNumber: 1, elements: [] },
		canvasSize: { width: 1280, height: 720 },
		scale: 1,
		mediaDataUrls: new Map<string, string>(),
		t: createTranslator(),
		smartArt3D: false,
		surfaceChart3D: false,
		barChart3D: false,
		lineChart3D: false,
		areaChart3D: false,
		pieChart3D: false,
		presenting: false,
		registry,
		renderElement(element, zIndex) {
			return registry.resolve(element.type)(element, zIndex, context);
		},
	};
	return context;
}

function bar3DPictureElement(): PptxElement {
	const chartData: PptxChartData = {
		chartType: 'bar3D',
		categories: ['Q1'],
		series: [
			{
				name: 'A',
				values: [10],
				dataPoints: [
					{
						idx: 0,
						picture: {
							imageUrl: IMAGE_URL,
							applyToFront: true,
							applyToSides: false,
							applyToEnd: false,
						},
					},
				],
			},
		],
	};
	return { type: 'chart', id: 'ch_bar3d_pic', x: 0, y: 0, width: 400, height: 300, chartData };
}

describe('renderChartElement - bar3D untargeted-face picture colour sample', () => {
	it('paints the untargeted side/end faces from the resolved point/series colour before any sample is cached', () => {
		const container = renderChartElement(bar3DPictureElement(), 1, buildContext()) as HTMLElement;
		const fills = [...container.querySelectorAll('polygon')].map((p) => p.getAttribute('fill'));
		expect(fills).not.toContain(resolveUntargetedBarFaceFill('side', SAMPLED_COLOR));
		expect(fills).not.toContain(resolveUntargetedBarFaceFill('end', SAMPLED_COLOR));
	});

	it('paints the untargeted side/end faces from the sampled colour once it is cached', async () => {
		ensureBarFacePicturePixelSampled(IMAGE_URL, vi.fn().mockResolvedValue(SAMPLED_COLOR));
		await vi.waitFor(() => {
			expect(getCachedBarFacePicturePixelColor(IMAGE_URL)).toBe(SAMPLED_COLOR);
		});

		const container = renderChartElement(bar3DPictureElement(), 1, buildContext()) as HTMLElement;
		const fills = [...container.querySelectorAll('polygon')].map((p) => p.getAttribute('fill'));
		expect(fills).toContain(resolveUntargetedBarFaceFill('side', SAMPLED_COLOR));
		expect(fills).toContain(resolveUntargetedBarFaceFill('end', SAMPLED_COLOR));
	});

	it('repaints an already-mounted chart when a real sample resolves', async () => {
		const container = renderChartElement(bar3DPictureElement(), 1, buildContext()) as HTMLElement;
		document.body.appendChild(container);
		try {
			const svg = container.querySelector('svg');
			ensureBarFacePicturePixelSampled('blob:live-chart-event', async () => SAMPLED_COLOR);
			await vi.waitFor(() => expect(container.querySelector('svg')).not.toBe(svg));
			const fills = [...container.querySelectorAll('polygon')].map((p) => p.getAttribute('fill'));
			expect(fills.length).toBeGreaterThanOrEqual(2);
		} finally {
			container.remove();
		}
	});

	it('does not repaint a detached chart when a real sample resolves', async () => {
		const container = renderChartElement(bar3DPictureElement(), 1, buildContext()) as HTMLElement;
		document.body.appendChild(container);
		container.remove();
		const svg = container.querySelector('svg');
		ensureBarFacePicturePixelSampled('blob:detached-chart-event', async () => SAMPLED_COLOR);
		await vi.waitFor(() =>
			expect(getCachedBarFacePicturePixelColor('blob:detached-chart-event')).toBe(SAMPLED_COLOR),
		);
		expect(container.querySelector('svg')).toBe(svg);
	});
});
