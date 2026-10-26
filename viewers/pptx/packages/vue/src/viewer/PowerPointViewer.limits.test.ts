import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	LOAD_LIMIT_DECK_SLIDES,
	loadLimitDeck,
	OVER_SLIDE_LIMIT_MESSAGE,
} from '../../../shared/test-utils/load-limits.mjs';
import PowerPointViewer from './PowerPointViewer.vue';

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
	vi.restoreAllMocks();
});

async function mountWithLimits(limits: { maxSlides: number } | undefined) {
	const content = await loadLimitDeck(PptxHandler);
	const load = vi.spyOn(PptxHandler.prototype, 'load');
	// No Google Fonts request from the test environment once a deck renders.
	mounted = mount(PowerPointViewer, { props: { content, limits, googleFonts: false } });
	await vi.waitFor(async () => {
		await flushPromises();
		expect(load).toHaveBeenCalledWith(expect.any(ArrayBuffer), expect.anything());
		expect(mounted?.find('.pptx-vue-loading').exists()).toBeFalsy();
	});
	return load;
}

describe('powerPointViewer limits prop', () => {
	it('passes the limits to PptxHandler.load and shows the localized limit error', async () => {
		const limits = { maxSlides: LOAD_LIMIT_DECK_SLIDES - 1 };
		const load = await mountWithLimits(limits);
		expect(load).toHaveBeenCalledWith(expect.any(ArrayBuffer), expect.objectContaining({ limits }));
		await vi.waitFor(async () => {
			await flushPromises();
			expect(mounted!.find('.pptx-vue-error').exists()).toBeTruthy();
		});
		expect(mounted!.find('.pptx-vue-error').text()).toContain('Failed to load presentation.');
		expect(mounted!.find('.pptx-vue-error-detail').text()).toBe(OVER_SLIDE_LIMIT_MESSAGE);
	});

	it('keeps an omitted limits prop undefined, so the core defaults apply', async () => {
		const load = await mountWithLimits(undefined);
		expect(load.mock.calls[0]?.[1]).toHaveProperty('limits', undefined);
		await flushPromises();
		expect(mounted!.find('.pptx-vue-error').exists()).toBeFalsy();
	});
});
