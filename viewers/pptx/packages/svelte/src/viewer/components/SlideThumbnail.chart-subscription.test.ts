/**
 * A thumbnail's chart subscribes to the document-wide picture colour sampler;
 * unmounting the thumbnail (a host strip dropping or replacing it) must retire
 * that subscription, not leave it for the next sample to find.
 */
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { describe, expect, it, vi } from 'vitest';

import { SlideThumbnail } from '../../index';

const sampler = vi.hoisted(() => ({ active: 0 }));

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		subscribeBarFacePicturePixelSamples: (listener: () => void) => {
			sampler.active += 1;
			const off = actual.subscribeBarFacePicturePixelSamples(listener);
			let open = true;
			return () => {
				if (open) {
					open = false;
					sampler.active -= 1;
				}
				off();
			};
		},
	};
});

const chart = {
	type: 'chart',
	id: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: { chartType: 'bar', categories: ['A'], series: [{ name: 'One', values: [1] }] },
} as unknown as PptxElement;
const slide = {
	id: 'ppt/slides/slide1.xml',
	rId: 'rId2',
	slideNumber: 1,
	elements: [chart],
} as unknown as PptxSlide;

describe('slideThumbnail chart subscription', () => {
	it('retires the chart sampler subscription when the thumbnail unmounts', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		const component = mount(SlideThumbnail, {
			target,
			props: { slide, canvasSize: { width: 960, height: 540 } },
		});
		flushSync();
		expect(target.querySelector('svg')).not.toBeNull();
		expect(sampler.active).toBe(1);
		void unmount(component);
		flushSync();
		expect(sampler.active).toBe(0);
		target.remove();
	});
});
