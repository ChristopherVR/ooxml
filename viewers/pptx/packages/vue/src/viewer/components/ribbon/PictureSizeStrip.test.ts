/**
 * Picture Format > Size > Crop: the contextual tab draws the shared Size
 * strip, and its Crop split button drives the real crop controller (the same
 * one Home > Arrange uses).
 */
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';
import { nextTick } from 'vue';

import { mountWithMergeCrop } from '../../composables/merge-crop-test-harness';
import type { MergeCropHarness } from '../../composables/merge-crop-test-harness';
import ContextualTabSection from './ContextualTabSection.vue';

const PICTURE = {
	id: 'pic',
	type: 'picture',
	x: 300,
	y: 0,
	width: 200,
	height: 100,
} as PptxElement;
const SHAPE = { id: 'a', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;

let harness: MergeCropHarness | null = null;
function mountTab(selected: string[], props: Record<string, unknown> = {}): MergeCropHarness {
	harness = mountWithMergeCrop([SHAPE, PICTURE], selected, ContextualTabSection, {
		tab: 'pictureFormat',
		...props,
	});
	return harness;
}
afterEach(() => {
	harness?.wrapper.unmount();
	harness = null;
});

const sizeControl = (h: MergeCropHarness, hook: string) =>
	h.wrapper.get<HTMLButtonElement>(
		`[data-ribbon-group="pictureFormat.size"] [data-pptx-ribbon-control="${hook}"]`,
	);

describe('picture Format > Size', () => {
	it('toggles crop mode and keeps its presses inside the crop session', async () => {
		const h = mountTab(['pic']);
		expect(
			h.wrapper
				.find('[data-pptx-crop-keep] [data-ribbon-control="pictureFormat.size.crop"]')
				.exists(),
		).toBeTruthy();
		expect(sizeControl(h, 'crop').attributes('aria-pressed')).toBe('false');
		await sizeControl(h, 'crop').trigger('click');
		expect(h.controller.cropActive.value).toBeTruthy();
		expect(sizeControl(h, 'crop').attributes('aria-pressed')).toBe('true');
		await sizeControl(h, 'crop').trigger('click');
		expect(h.controller.cropActive.value).toBeFalsy();
	});

	it('crops to an aspect ratio from the caret menu', async () => {
		const h = mountTab(['pic']);
		await sizeControl(h, 'crop-menu').trigger('click');
		await h.wrapper.get('[data-pptx-crop-aspect="1:1"]').trigger('click');
		expect(h.elements().find((el) => el.id === 'pic')?.width).toBeCloseTo(100);
	});

	it('is disabled without a picture and hidden when the host hides crop', async () => {
		const h = mountTab(['a']);
		expect(sizeControl(h, 'crop').attributes('disabled')).toBeDefined();
		h.wrapper.unmount();
		const hidden = mountTab(['pic'], { hiddenActions: ['crop'] });
		await nextTick();
		expect(sizeControl(hidden, 'crop').element.hidden).toBeTruthy();
		expect(sizeControl(hidden, 'crop-menu').element.hidden).toBeTruthy();
	});
});
