/**
 * `SlidesPane` inside `PowerPointViewer`: the full viewer renders the same
 * rail, so its thumbnails resolve fields from the same deck-wide context the
 * canvas does, and selecting a thumbnail only navigates.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type { FieldSubstitutionContext } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ComponentPublicInstance, MaybeRefOrGetter } from 'vue';
import { toValue } from 'vue';

import { FieldContextKey } from '../composables/field-context';
import PowerPointViewer from '../PowerPointViewer.vue';
import type { PowerPointViewerExpose } from '../types';
import SlidesPane from './SlidesPane.vue';

// There is no IndexedDB here, and the recovery write is fire-and-forget.
// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	saveAutosaveSnapshot: () => Promise.resolve(true),
}));

const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx');

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
});

async function mountViewer(canEdit: boolean) {
	const wrapper = mount(PowerPointViewer, {
		props: { content: new Uint8Array(readFileSync(FIXTURE)), canEdit },
	});
	mounted = wrapper;
	const api = wrapper.vm as unknown as PowerPointViewerExpose;
	await vi.waitFor(
		async () => {
			await flushPromises();
			expect(api.getSlideCount()).toBe(3);
			expect(wrapper.findAll('[data-pptx-chrome="slide-frame"]')).toHaveLength(3);
		},
		{ timeout: 15000 },
	);
	return wrapper;
}

/** The field context a component provides to its subtree. */
function providedFieldContext(vm: ComponentPublicInstance): FieldSubstitutionContext | undefined {
	const provides = (vm.$ as unknown as { provides: Record<symbol, unknown> }).provides;
	const source = provides[FieldContextKey as symbol] as
		| MaybeRefOrGetter<FieldSubstitutionContext | undefined>
		| undefined;
	return source ? toValue(source) : undefined;
}

describe('slidesPane in PowerPointViewer', () => {
	it("hands thumbnails the canvas's deck-wide field context", async () => {
		const wrapper = await mountViewer(false);
		const canvas = providedFieldContext(wrapper.vm);
		const pane = providedFieldContext(wrapper.findComponent(SlidesPane).vm);
		expect(canvas?.firstSlideNumber).toBe(45);
		// The same deck-wide fields; each thumbnail supplies its own number and title.
		expect(pane).toStrictEqual({ ...canvas, slideNumber: undefined, slideTitle: undefined });
		const frames = wrapper.findAll('[data-pptx-chrome="slide-frame"]').map((f) => f.text());
		expect(frames).toStrictEqual([
			expect.stringContaining('Slide 45'),
			expect.stringContaining('Slide 46'),
			expect.stringContaining('Slide 47'),
		]);
	});

	it('selecting a thumbnail moves the canvas without marking the deck dirty', async () => {
		const wrapper = await mountViewer(true);
		const api = wrapper.vm as unknown as PowerPointViewerExpose;
		// Loading settles the flag back to clean; only what the click emits counts.
		const dirtyBefore = wrapper.emitted('dirty-change')?.length ?? 0;
		expect(wrapper.emitted('dirty-change')?.at(-1)).toStrictEqual([false]);
		await wrapper.findAll('[data-pptx-chrome="slide-row"]')[2].trigger('click');
		await flushPromises();
		expect(api.getActiveSlideIndex()).toBe(2);
		expect(wrapper.emitted('active-slide-change')?.at(-1)).toStrictEqual([2]);
		expect(wrapper.emitted('dirty-change')?.length ?? 0).toBe(dirtyBefore);
		expect(wrapper.emitted('change')).toBeUndefined();
	});
});
