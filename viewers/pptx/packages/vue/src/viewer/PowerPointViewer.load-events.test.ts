import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';

/** `load` / `error` emits are wired to the real load pipeline. */
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

async function settle(done: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !done(); attempt += 1) {
		await flushPromises();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
}

describe('load and error emits', () => {
	it('emits load with the slide count and canvas size once a deck has loaded', async () => {
		const wrapper = mount(PowerPointViewer, { props: { content: await sampleDeck(3) } });
		mounted = wrapper;
		await settle(() => (wrapper.emitted('load')?.length ?? 0) > 0);

		expect(wrapper.emitted('load')).toHaveLength(1);
		const detail = wrapper.emitted('load')?.[0]?.[0] as {
			slideCount: number;
			canvasSize: { width: number };
		};
		expect(detail.slideCount).toBe(3);
		expect(detail.canvasSize.width).toBeGreaterThan(0);
		expect(wrapper.emitted('error')).toBeUndefined();
	});

	it('emits error, and no load, when the bytes are not a presentation', async () => {
		const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
		try {
			const wrapper = mount(PowerPointViewer, { props: { content: new Uint8Array([1, 2, 3, 4]) } });
			mounted = wrapper;
			await settle(() => (wrapper.emitted('error')?.length ?? 0) > 0);

			expect(wrapper.emitted('error')).toHaveLength(1);
			expect(typeof wrapper.emitted('error')?.[0]?.[0]).toBe('string');
			expect(wrapper.emitted('load')).toBeUndefined();
		} finally {
			errorSpy.mockRestore();
		}
	});
});
