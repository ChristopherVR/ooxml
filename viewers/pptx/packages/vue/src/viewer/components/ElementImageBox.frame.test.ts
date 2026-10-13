import { mount } from '@vue/test-utils';
import type { ImagePptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import ElementImageBox from './ElementImageBox.vue';

/** A Picture Styles frame: a centred 8pt red line and an outer shadow. */
function framedPicture(withFrame: boolean): ImagePptxElement {
	return {
		type: 'image',
		id: 'pic-frame',
		x: 0,
		y: 0,
		width: 200,
		height: 150,
		imageData: 'data:image/png;base64,AA==',
		...(withFrame
			? {
					shapeStyle: {
						strokeWidth: 10.67,
						strokeFillMode: 'solid',
						strokeColor: '#FF0000',
						shadowColor: '#000000',
						shadowOpacity: 0.6,
						shadowBlur: 5.33,
						shadowAngle: 45,
						shadowDistance: 10.67,
					},
				}
			: {}),
	};
}

function mountBox(element: ImagePptxElement) {
	return mount(ElementImageBox, { props: { element, mediaDataUrls: new Map(), zIndex: 0 } });
}

describe('element image box picture frame (a:ln + a:outerShdw)', () => {
	it('paints the whole frame and the shadow from an unclipped box', () => {
		const box = mountBox(framedPicture(true)).get('.pptx-vue-image').element as HTMLElement;
		expect(box.style.overflow).toBe('visible');
		expect(box.style.filter).toMatch(/drop-shadow\(/);
		const surface = box.querySelector<HTMLElement>('.pptx-vue-image-surface');
		expect(surface?.style.overflow).toBe('hidden');
		expect(surface?.querySelector('img')).not.toBeNull();
		const path = [...box.querySelectorAll('svg path')].find(
			(p) => p.getAttribute('stroke') === '#FF0000',
		);
		expect(path).toBeDefined();
		expect(
			box.querySelector('img')!.compareDocumentPosition(path!) & Node.DOCUMENT_POSITION_FOLLOWING,
		).toBeTruthy();
		// The outline is not inside the clipped surface, so its outer half shows.
		expect(surface?.contains(path ?? null)).toBe(false);
	});

	it('keeps the clip on the box for a picture without a frame', () => {
		const box = mountBox(framedPicture(false)).get('.pptx-vue-image').element as HTMLElement;
		expect(box.style.overflow).toBe('hidden');
		expect(box.style.filter).toBe('');
		expect(box.querySelector<HTMLElement>('.pptx-vue-image-surface')?.style.display).toBe(
			'contents',
		);
	});
});
