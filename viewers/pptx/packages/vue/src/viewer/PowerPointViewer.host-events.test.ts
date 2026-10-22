// @vitest-environment happy-dom
/**
 * The host events every binding reports the same way (issue 151, item 5), driven through the
 * mounted component: `change`, `presentation-change`, `autosave`, `autosave-status` and
 * `autosave-toggle`.
 */
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

// There is no IndexedDB here, and the recovery write is fire-and-forget.
// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	saveAutosaveSnapshot: () => Promise.resolve(true),
}));

async function sampleDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
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

async function mountLoaded(props: Record<string, unknown> = {}) {
	const wrapper = mount(PowerPointViewer, {
		props: { content: await sampleDeck(), canEdit: true, filePath: 'events.pptx', ...props },
	});
	mounted = wrapper;
	await settle(() => (wrapper.emitted('load')?.length ?? 0) > 0);
	await flushPromises();
	return { wrapper, viewer: wrapper.vm as unknown as PowerPointViewerExpose };
}

describe('host events through the mounted component', () => {
	it('emits change for a committed edit, not for the load', async () => {
		const { wrapper, viewer } = await mountLoaded();
		expect(wrapper.emitted('change')).toBeUndefined();
		viewer.addSlide();
		await flushPromises();
		expect(wrapper.emitted('change')?.length).toBeGreaterThan(0);
	});

	it('emits dirty-change when an edit raises the unsaved-changes flag', async () => {
		const { wrapper, viewer } = await mountLoaded();
		expect(wrapper.emitted('dirty-change')?.at(-1)).toStrictEqual([false]);
		viewer.addSlide();
		await flushPromises();
		expect(wrapper.emitted('dirty-change')?.at(-1)).toStrictEqual([true]);
	});

	it('emits presentation-change when the slide show starts and stops', async () => {
		const { wrapper, viewer } = await mountLoaded();
		expect(wrapper.emitted('presentation-change')).toBeUndefined();
		viewer.setMode('present');
		await flushPromises();
		expect(wrapper.emitted('presentation-change')?.at(-1)).toStrictEqual([true]);
		viewer.setMode('edit');
		await flushPromises();
		expect(wrapper.emitted('presentation-change')?.at(-1)).toStrictEqual([false]);
	});

	it('emits autosave with the snapshot bytes and autosave-status through the lifecycle', async () => {
		const { wrapper, viewer } = await mountLoaded({ autosaveIntervalMs: 30 });
		viewer.addSlide();
		await settle(() => (wrapper.emitted('autosave')?.length ?? 0) > 0);
		expect(wrapper.emitted('autosave')?.[0]?.[0]).toBeInstanceOf(Uint8Array);
		const statuses = (wrapper.emitted('autosave-status') ?? []).map((args) => args[0]);
		expect(statuses).toContain('saving');
		expect(statuses).toContain('saved');
	});

	it('emits autosave-toggle with the new preference when the title-bar switch is used', async () => {
		const { wrapper } = await mountLoaded();
		const titleBar = wrapper.find('pptx-ui-title-bar');
		expect(titleBar.exists()).toBeTruthy();
		titleBar.element.dispatchEvent(new CustomEvent('toggle-autosave', { bubbles: true }));
		await flushPromises();
		expect(wrapper.emitted('autosave-toggle')).toStrictEqual([[false]]);
		titleBar.element.dispatchEvent(new CustomEvent('toggle-autosave', { bubbles: true }));
		await flushPromises();
		expect(wrapper.emitted('autosave-toggle')?.at(-1)).toStrictEqual([true]);
	});

	it('does not emit autosave-toggle when the host forbade autosave', async () => {
		const { wrapper } = await mountLoaded({ autosave: false });
		wrapper
			.find('pptx-ui-title-bar')
			.element.dispatchEvent(new CustomEvent('toggle-autosave', { bubbles: true }));
		await flushPromises();
		expect(wrapper.emitted('autosave-toggle')).toBeUndefined();
	});
});
