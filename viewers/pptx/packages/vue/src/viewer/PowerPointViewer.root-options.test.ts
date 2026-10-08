import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose, PowerPointViewerProps } from './types';

/**
 * The root display props `showToolbar`, `showThumbnails` and `initialSlide`,
 * wired through the whole viewer. Their defaults and clamping come from the
 * shared `resolveViewerRootOptions` / `resolveInitialSlideIndex`
 * (ooxml-ui/pptx), so they behave as they do in every other binding.
 */
async function sampleDeck(slideCount: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slideCount });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
});

async function mountViewer(
	props: Partial<PowerPointViewerProps>,
	slideCount = 1,
): Promise<{ wrapper: VueWrapper; viewer: PowerPointViewerExpose }> {
	const content = await sampleDeck(slideCount);
	const wrapper = mount(PowerPointViewer, { props: { content, canEdit: true, ...props } });
	mounted = wrapper;
	const viewer = wrapper.vm as unknown as PowerPointViewerExpose;
	for (let attempt = 0; attempt < 100 && viewer.getSlideCount() === 0; attempt += 1) {
		await flushPromises();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
	await flushPromises();
	expect(viewer.getSlideCount()).toBe(slideCount);
	return { wrapper, viewer };
}

const has = (wrapper: VueWrapper, selector: string): boolean => wrapper.find(selector).exists();
const RIBBON = '[data-pptx-chrome="ribbon"]';
const SLIDES = '[data-pptx-chrome="slides"]';
const STATUS_BAR = 'pptx-ui-status-bar';

describe('powerPointViewer root display props', () => {
	it('shows the toolbar, status bar and thumbnails and opens slide 0 by default', async () => {
		const { wrapper, viewer } = await mountViewer({}, 3);
		expect(has(wrapper, RIBBON)).toBeTruthy();
		expect(has(wrapper, STATUS_BAR)).toBeTruthy();
		expect(has(wrapper, SLIDES)).toBeTruthy();
		expect(viewer.getActiveSlideIndex()).toBe(0);
	});

	it('hides the ribbon and status bar when showToolbar is false', async () => {
		const { wrapper } = await mountViewer({ showToolbar: false });
		expect(has(wrapper, RIBBON)).toBeFalsy();
		expect(has(wrapper, STATUS_BAR)).toBeFalsy();
		expect(has(wrapper, SLIDES)).toBeTruthy();
	});

	it('hides the slide thumbnail pane when showThumbnails is false', async () => {
		const { wrapper } = await mountViewer({ showThumbnails: false });
		expect(has(wrapper, SLIDES)).toBeFalsy();
		expect(has(wrapper, RIBBON)).toBeTruthy();
	});

	it('opens the deck on initialSlide', async () => {
		const { viewer } = await mountViewer({ initialSlide: 2 }, 3);
		expect(viewer.getActiveSlideIndex()).toBe(2);
	});

	it('clamps an out-of-range initialSlide into the deck', async () => {
		const { viewer } = await mountViewer({ initialSlide: 99 }, 3);
		expect(viewer.getActiveSlideIndex()).toBe(2);
	});
});
