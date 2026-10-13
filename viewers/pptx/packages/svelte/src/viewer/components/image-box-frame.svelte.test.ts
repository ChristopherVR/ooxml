/**
 * A picture's own frame (`a:ln`) and outer shadow must show in full: Svelte
 * stroked the line inside the box that clips the picture, so only its inner
 * half showed, and it never applied the shadow.
 */
import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import ImageBox from './ImageBox.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function mountBox(withFrame: boolean): HTMLElement {
	const element = {
		type: 'picture',
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
	} as unknown as PptxElement;
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(ImageBox, {
		target,
		props: { element, mediaDataUrls: new Map<string, string>(), zIndex: 1 },
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target.querySelector<HTMLElement>('.pptx-svelte-image') as HTMLElement;
}

describe('image box picture frame (a:ln + a:outerShdw)', () => {
	it('paints the whole frame and the shadow from an unclipped box', () => {
		const box = mountBox(true);
		expect(box.style.overflow).toBe('visible');
		expect(box.style.filter).toMatch(/drop-shadow\(/);
		const surface = box.querySelector<HTMLElement>(':scope > .pptx-svelte-image-surface');
		expect(surface?.style.overflow).toBe('hidden');
		expect(surface?.querySelector('img')).not.toBeNull();
		const path = [...box.querySelectorAll('svg path')].find(
			(p) => p.getAttribute('stroke') === '#FF0000',
		);
		expect(path).toBeDefined();
		expect(
			box.querySelector('img')!.compareDocumentPosition(path!) & Node.DOCUMENT_POSITION_FOLLOWING,
		).toBeTruthy();
		expect(surface?.contains(path ?? null)).toBe(false);
	});

	it('keeps the single clipped box for a picture without a frame', () => {
		const box = mountBox(false);
		expect(box.style.overflow).toBe('hidden');
		expect(box.querySelector('.pptx-svelte-image-surface')).toBeNull();
		expect(box.querySelector(':scope > img')).not.toBeNull();
	});
});
