import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose, PowerPointViewerProps } from './types';

/**
 * The root display props `showToolbar`, `showThumbnails`,
 * `showCompatibilityToasts` and `initialSlide`, wired through the whole
 * viewer. Their defaults and clamping come from the shared `resolveViewerRootOptions` / `resolveInitialSlideIndex`
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

/** A deck whose presentation part carries unmodelled markup, so it loads with a compatibility warning. */
async function warnedDeck(): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await sampleDeck(1));
	const xml = await zip.file('ppt/presentation.xml')?.async('string');
	if (!xml) {
		throw new Error('the sample deck has no presentation part');
	}
	zip.file(
		'ppt/presentation.xml',
		xml.replace('</p:presentation>', '<p:unmodelledMarker/></p:presentation>'),
	);
	return zip.generateAsync({ type: 'uint8array' });
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
	deck?: Uint8Array,
): Promise<{ wrapper: VueWrapper; viewer: PowerPointViewerExpose }> {
	const content = deck ?? (await sampleDeck(slideCount));
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
const COMPAT_TOASTS = 'pptx-ui-compat-toasts';

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

	it('shows the compatibility toast stack by default when the deck has warnings', async () => {
		const { wrapper } = await mountViewer({}, 1, await warnedDeck());
		expect(has(wrapper, COMPAT_TOASTS)).toBeTruthy();
	});

	it('hides the compatibility toast stack when showCompatibilityToasts is false', async () => {
		const { wrapper } = await mountViewer(
			{ showCompatibilityToasts: false },
			1,
			await warnedDeck(),
		);
		expect(has(wrapper, COMPAT_TOASTS)).toBeFalsy();
		expect(has(wrapper, RIBBON)).toBeTruthy();
	});
});
