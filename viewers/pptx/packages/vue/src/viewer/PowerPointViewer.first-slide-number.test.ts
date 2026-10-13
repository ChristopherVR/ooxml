/**
 * `p:presentation/@firstSlideNum="45"`: PowerPoint numbers slide `i` as
 * 45 + i, so the `slidenum` field on slide 1 reads 45. The viewer showed the
 * slide's position (1) instead.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushPromises, mount } from '@vue/test-utils';
import { expect, test, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

test('numbers slide-number fields from the deck start number', async () => {
	const content = new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx')),
	);
	const wrapper = mount(PowerPointViewer, { props: { content, canEdit: false } });
	try {
		const api = wrapper.vm as unknown as PowerPointViewerExpose;
		await vi.waitFor(
			async () => {
				await flushPromises();
				expect(api.getSlideCount()).toBe(3);
			},
			{ timeout: 15000 },
		);
		const stage = wrapper.get('[data-pptx-viewport]').text();
		expect(stage).toContain('Slide 45');
		expect(stage).not.toContain('Slide 1');
	} finally {
		wrapper.unmount();
	}
});
