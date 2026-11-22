// @vitest-environment jsdom
import type { SummaryZoomTarget } from 'ooxml-core/pptx';
import { afterEach, beforeAll, expect, it } from 'vitest';

import { definePptxZoomCover, type ZoomCoverPatchEvent } from './zoom-cover-control';

beforeAll(() => definePptxZoomCover(customElements));
afterEach(() => document.body.replaceChildren());

it('resets the chosen Summary Zoom cover without editing another tile', () => {
	const frame = { x: 10, y: 10, width: 50, height: 30 };
	const first: SummaryZoomTarget = {
		...frame,
		sectionId: 'first',
		title: 'Introduction',
		targetSlideIndex: 1,
	};
	const second: SummaryZoomTarget = {
		...frame,
		sectionId: 'second',
		title: 'Details',
		targetSlideIndex: 2,
		zoomImageType: 'cover',
		imagePath: 'ppt/media/details.png',
		imageData: 'cover',
	};
	const cover = document.createElement('pptx-ui-zoom-cover');
	cover.state = {
		element: {
			...frame,
			id: 'summary',
			type: 'zoom',
			zoomType: 'summary',
			targetSlideIndex: 1,
			summaryTargets: [first, second],
		},
		translate: (key) => key,
	};
	const intents: ZoomCoverPatchEvent['detail'][] = [];
	cover.addEventListener('zoom-cover-patch', (event) =>
		intents.push((event as ZoomCoverPatchEvent).detail),
	);
	document.body.append(cover);

	const reset = Array.from(cover.querySelectorAll('button')).find(
		(button) => button.textContent === 'pptx.graphic.zoomResetCover',
	);
	expect(reset?.disabled).toBe(true);
	const tiles = cover.querySelector('pptx-ui-select');
	expect(tiles?.options.map((option) => option.textContent)).toEqual(['Introduction', 'Details']);
	tiles!.value = second.sectionId;
	tiles!.dispatchEvent(new Event('change'));
	expect(reset?.disabled).toBe(false);
	reset?.click();

	expect(intents).toEqual([
		{
			summaryTargets: [
				first,
				{ ...second, zoomImageType: 'preview', imagePath: undefined, imageData: undefined },
			],
		},
	]);
	expect(second.zoomImageType).toBe('cover');
});
