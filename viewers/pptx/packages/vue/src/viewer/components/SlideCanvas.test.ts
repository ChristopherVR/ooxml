import { mount } from '@vue/test-utils';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import Viewer from '../PowerPointViewer.vue';
import type { CanvasSize } from '../types';
import SlideCanvas from './SlideCanvas.vue';

const canvasSize: CanvasSize = { width: 960, height: 540 };

afterEach(() => {
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

function makeSlide(overrides: Partial<PptxSlide>): PptxSlide {
	return { id: 'slide-1', elements: [], ...overrides } as unknown as PptxSlide;
}

function mountCanvas(slide: PptxSlide | undefined) {
	return mount(SlideCanvas, {
		props: { slide, canvasSize, mediaDataUrls: new Map<string, string>() },
	});
}

describe('slideCanvas slide-region contract', () => {
	it('exposes the actual scaled stage as a public coordinate origin', async () => {
		const wrapper = mountCanvas(makeSlide({}));
		const stage = wrapper.vm.getStageElement();
		expect(stage).toBe(wrapper.get('.pptx-vue-stage').element);
		await wrapper.setProps({ zoom: 1.5 });
		expect(wrapper.vm.getStageElement()).toBe(stage);
		expect(stage?.style.transform).toContain('scale(1.5)');
		wrapper.unmount();
		expect(wrapper.vm.getStageElement()).toBeNull();
	});

	// The e2e accessibility contract is ONE visible slide region for the
	// current slide: the wrapper. The inner interactive SlideStage must not
	// duplicate the aria-roledescription (a strict Playwright locator on
	// '[aria-roledescription="slide"]' resolved to 2 elements before this).
	it('renders exactly one aria-roledescription="slide" region', () => {
		const wrapper = mountCanvas(makeSlide({}));
		expect(wrapper.findAll('[aria-roledescription="slide"]')).toHaveLength(1);
		const region = wrapper.get('[aria-roledescription="slide"]');
		expect(region.classes()).toContain('pptx-vue-canvas-wrapper');
	});
});

describe('slideCanvas background inheritance', () => {
	// The e2e contract reads the background off the aria-roledescription="slide"
	// region (the wrapper), mirroring React/Angular. A slide whose background is
	// inherited from its master/layout arrives with `backgroundColor` already
	// resolved by core, so the labelled region must paint it (not stay
	// transparent).
	it('paints the resolved slide background on the slide region', () => {
		const wrapper = mountCanvas(makeSlide({ backgroundColor: '#000000' }));
		const region = wrapper.get('[aria-roledescription="slide"]');
		expect(region.attributes('style')).toContain('background-color: #000000');
	});

	it('falls back to white when no background is resolved', () => {
		const wrapper = mountCanvas(makeSlide({}));
		const region = wrapper.get('[aria-roledescription="slide"]');
		expect(region.attributes('style')).toContain('background-color: #ffffff');
	});

	it('treats an explicit "transparent" background as the white fallback', () => {
		const wrapper = mountCanvas(makeSlide({ backgroundColor: 'transparent' }));
		const region = wrapper.get('[aria-roledescription="slide"]');
		expect(region.attributes('style')).toContain('background-color: #ffffff');
	});
});

describe('slideCanvas viewport fit policy', () => {
	it('declares host fit options as runtime props on the public viewer', () => {
		expect(Viewer.props).toHaveProperty('fitPadding');
		expect(Viewer.props).toHaveProperty('maxFitScale');
	});

	function measuredCanvas(width = 960, height = 540) {
		vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockReturnValue(width);
		vi.spyOn(HTMLElement.prototype, 'clientHeight', 'get').mockReturnValue(height);
		return mountCanvas(makeSlide({}));
	}

	it('preserves default padding and responds to explicit padding changes', async () => {
		const wrapper = measuredCanvas();
		const fit = () => wrapper.emitted('update:fitScale')!.at(-1)![0];
		expect(fit()).toBeCloseTo(508 / 540);
		await wrapper.setProps({ fitPadding: 0 });
		expect(fit()).toBe(1);
		await wrapper.setProps({ fitPadding: { horizontal: 80, vertical: 0 } });
		expect(fit()).toBeCloseTo(800 / 960);
		wrapper.unmount();
	});

	it('can opt into enlargement or a finite cap without changing supplied zoom', async () => {
		const wrapper = measuredCanvas(1920, 1080);
		const fit = () => wrapper.emitted('update:fitScale')!.at(-1)![0];
		expect(fit()).toBe(1);
		await wrapper.setProps({ fitPadding: 0, maxFitScale: null, zoom: 3 });
		expect(fit()).toBe(2);
		expect(wrapper.get('[aria-roledescription="slide"]').attributes('style')).toContain(
			'width: 2880px',
		);
		await wrapper.setProps({ maxFitScale: 1.5 });
		expect(fit()).toBe(1.5);
		await wrapper.setProps({ maxFitScale: -1, fitPadding: -1 });
		expect(fit()).toBe(1);
		wrapper.unmount();
	});

	it('retains ruler gutters in default and zero-padding modes', async () => {
		const wrapper = measuredCanvas();
		const fit = () => wrapper.emitted('update:fitScale')!.at(-1)![0];
		await wrapper.setProps({ showRulers: true });
		expect(fit()).toBeCloseTo(488 / 540);
		await wrapper.setProps({ fitPadding: 0 });
		expect(fit()).toBeCloseTo(920 / 960);
		await wrapper.setProps({ showRulers: false });
		expect(fit()).toBe(1);
		wrapper.unmount();
	});

	it.each([
		[801.5, 700],
		[1200, 451.5],
	])('fits a %spx x %spx viewport without exceeding it', async (width, height) => {
		// What a browser reports for a box laid out at a fractional size: whole-pixel client and
		// offset sizes, and the unrounded rect.
		vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockReturnValue(Math.round(width));
		vi.spyOn(HTMLElement.prototype, 'clientHeight', 'get').mockReturnValue(Math.round(height));
		vi.spyOn(HTMLElement.prototype, 'offsetWidth', 'get').mockReturnValue(Math.round(width));
		vi.spyOn(HTMLElement.prototype, 'offsetHeight', 'get').mockReturnValue(Math.round(height));
		vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockReturnValue(
			new DOMRect(0, 0, width, height),
		);
		const wrapper = mount(SlideCanvas, {
			props: {
				slide: makeSlide({}),
				canvasSize,
				mediaDataUrls: new Map<string, string>(),
				fitPadding: 0,
				maxFitScale: null,
			},
		});
		const fit = wrapper.emitted('update:fitScale')!.at(-1)![0] as number;
		expect(canvasSize.width * fit).toBeLessThanOrEqual(width + 1e-9);
		expect(canvasSize.height * fit).toBeLessThanOrEqual(height + 1e-9);
		expect(fit).toBeCloseTo(Math.min(width / canvasSize.width, height / canvasSize.height), 9);
		wrapper.unmount();
	});

	it('refits on a padding-only change, which only the border box shows', async () => {
		const observers: { box: string | undefined; callback: () => void }[] = [];
		vi.stubGlobal(
			'ResizeObserver',
			class {
				private readonly callback: () => void;
				constructor(callback: () => void) {
					this.callback = callback;
				}
				observe(_node: Element, options?: ResizeObserverOptions) {
					observers.push({ box: options?.box, callback: this.callback });
				}
				disconnect() {}
			},
		);
		// A content-box viewport losing its padding: the client size drops, the content box
		// a default ResizeObserver watches stays put.
		const width = vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockReturnValue(1008);
		const height = vi.spyOn(HTMLElement.prototype, 'clientHeight', 'get').mockReturnValue(588);
		const wrapper = mountCanvas(makeSlide({}));
		await wrapper.setProps({ fitPadding: 0, maxFitScale: null });
		expect(wrapper.emitted('update:fitScale')!.at(-1)![0]).toBe(1.05);
		width.mockReturnValue(960);
		height.mockReturnValue(540);
		observers.find((observer) => observer.box === 'border-box')!.callback();
		expect(wrapper.emitted('update:fitScale')!.at(-1)![0]).toBe(1);
		wrapper.unmount();
	});

	it('remeasures fit on resize without altering authored canvas dimensions', async () => {
		let resize = () => {};
		vi.stubGlobal(
			'ResizeObserver',
			class {
				constructor(callback: () => void) {
					resize = callback;
				}
				observe() {}
				disconnect() {}
			},
		);
		const width = vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockReturnValue(960);
		const height = vi.spyOn(HTMLElement.prototype, 'clientHeight', 'get').mockReturnValue(540);
		const wrapper = mountCanvas(makeSlide({}));
		await wrapper.setProps({ fitPadding: 0, maxFitScale: null });
		width.mockReturnValue(1920);
		height.mockReturnValue(1080);
		resize();
		expect(wrapper.emitted('update:fitScale')!.at(-1)![0]).toBe(2);
		expect(wrapper.props('canvasSize')).toStrictEqual({ width: 960, height: 540 });
		wrapper.unmount();
	});
});
