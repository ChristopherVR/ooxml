import { describe, expect, it, vi } from 'vitest';

import { mayUseOriginBlobUrl, pictureUrl } from './picture-url';

const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');

describe('pictureUrl', () => {
	it('never gives an SVG picture a URL of the page origin', () => {
		const objectUrl = vi.fn(() => 'blob:https://app.example/1');
		const url = pictureUrl(svg, 'image/svg+xml', objectUrl);
		expect(objectUrl).not.toHaveBeenCalled();
		expect(url.startsWith('data:image/svg+xml;base64,')).toBe(true);
		// The same bytes, so the picture renders exactly as before.
		const payload = atob(url.slice(url.indexOf(',') + 1));
		expect(Uint8Array.from(payload, (c) => c.charCodeAt(0))).toEqual(svg);
	});

	it('keeps blob URLs for inert raster pictures, and data URLs without a factory', () => {
		const objectUrl = vi.fn(() => 'blob:https://app.example/2');
		expect(pictureUrl(new Uint8Array([1, 2]), 'IMAGE/PNG', objectUrl)).toBe(
			'blob:https://app.example/2',
		);
		expect(pictureUrl(new Uint8Array([1, 2]), 'image/png', null)).toBe(
			'data:image/png;base64,AQI=',
		);
	});

	it('treats every type that is not an inert raster image as needing an opaque origin', () => {
		for (const type of ['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp'])
			expect(mayUseOriginBlobUrl(type)).toBe(true);
		for (const type of ['image/svg+xml', 'text/html', 'application/pdf', 'image/x-emf', 'nonsense'])
			expect(mayUseOriginBlobUrl(type)).toBe(false);
	});
});
