// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it } from 'vitest';
import { createTranslator } from '../i18n/translator';
import { ensureBarFacePicturePixelSampled } from '../render/chart-bar3d-face-picture-sample';
import { releasePptxRenderCaches } from '../render/render-cache-lifecycle';
import { createDefaultRegistry } from './elements';
import { reRenderPresentationElements } from './presentation-rerender';
import { disposePptxRenderStage, renderSlideStage } from './slide-stage';
import type { ElementRenderContext } from './types';

const chart: PptxElement = {
	type: 'chart',
	id: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: { chartType: 'bar', categories: ['A'], series: [{ name: 'One', values: [1] }] },
};
const slide: PptxSlide = { id: 'slide', rId: 'rId1', slideNumber: 1, elements: [chart] };
const mounted = new Set<HTMLElement>();
const options = () => ({
	document,
	slide,
	canvasSize: { width: 800, height: 600 },
	mediaDataUrls: new Map<string, string>(),
	registry: createDefaultRegistry(),
	t: createTranslator(),
});

let sampleIndex = 0;
async function sample(): Promise<void> {
	ensureBarFacePicturePixelSampled(`blob:${sampleIndex++}`, async () => '#123456');
	// The sampler promise includes then/catch/finally; flush its complete chain.
	for (let i = 0; i < 6; i++) await Promise.resolve();
}
function mount(): HTMLElement {
	const stage = renderSlideStage(options());
	mounted.add(stage);
	document.body.appendChild(stage);
	return stage;
}
afterEach(() => {
	for (const stage of mounted) {
		disposePptxRenderStage(stage);
		stage.remove();
	}
	mounted.clear();
	releasePptxRenderCaches();
});

describe('slide stage disposal', () => {
	it('unsubscribes immediately, keeps other live charts subscribed, and is idempotent', async () => {
		const closed = mount(),
			live = mount();
		const closedSvg = closed.querySelector('svg'),
			liveSvg = live.querySelector('svg');
		disposePptxRenderStage(closed);
		disposePptxRenderStage(closed);
		await sample();
		expect(closed.querySelector('svg')).toBe(closedSvg);
		expect(live.querySelector('svg')).not.toBe(liveSvg);
	});

	it('cleans up charts constructed inside a group when a later renderer throws', async () => {
		const opts = options();
		let orphan: HTMLElement | null = null;
		const renderChart = opts.registry.resolve('chart');
		opts.registry.register('chart', (...args) => {
			orphan = renderChart(...args) as HTMLElement;
			return orphan;
		});
		opts.registry.register('shape', () => {
			throw new Error('host renderer failed');
		});
		opts.slide = {
			...slide,
			elements: [
				{
					type: 'group',
					id: 'group',
					x: 0,
					y: 0,
					width: 800,
					height: 600,
					children: [chart, { type: 'shape', id: 'error', x: 0, y: 0, width: 1, height: 1 }],
				},
			],
		};
		expect(() => renderSlideStage(opts)).toThrow('host renderer failed');
		expect(orphan).not.toBeNull();
		const node = orphan! as HTMLElement;
		mounted.add(node);
		document.body.appendChild(node);
		const svg = node.querySelector('svg');
		await sample();
		expect(node.querySelector('svg')).toBe(svg);
	});

	it('transfers chart repaint ownership when presentation builds preserve the wrapper', async () => {
		let context!: ElementRenderContext;
		const stage = renderSlideStage({
			...options(),
			captureContext: (value) => {
				context = value;
			},
		});
		mounted.add(stage);
		document.body.appendChild(stage);
		const wrapper = stage.querySelector('[data-element-id="chart"]');
		for (let frame = 0; frame < 20; frame++) {
			reRenderPresentationElements(context, stage, ['chart']);
		}
		const svg = stage.querySelector('svg');
		await sample();
		expect(stage.querySelector('[data-element-id="chart"]')).toBe(wrapper);
		expect(stage.querySelector('svg')).not.toBe(svg);
		disposePptxRenderStage(stage);
		const disposedSvg = stage.querySelector('svg');
		await sample();
		expect(stage.querySelector('svg')).toBe(disposedSvg);
	});
});
