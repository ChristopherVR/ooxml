import { afterEach, describe, expect, it, vi } from 'vitest';

import type { PptxSvgRasterizer } from '../../core/types';
import { rasterizeSvg, svgRasterSize } from './svg-rasterize';

const SVG = new TextEncoder().encode(
	'<svg xmlns="http://www.w3.org/2000/svg" width="100" height="60"/>',
);
const DOM_PNG = new Uint8Array([1, 2, 3]);

/** A browser-shaped runtime whose canvas records the size it was given. */
function stubBrowserCanvas(): { size: { width: number; height: number } } {
	const drawn = { size: { width: 0, height: 0 } };
	class FakeImage {
		naturalWidth = 100;
		naturalHeight = 60;
		onload: (() => void) | null = null;
		onerror: (() => void) | null = null;
		#src = '';
		get src(): string {
			return this.#src;
		}
		set src(url: string) {
			this.#src = url;
			queueMicrotask(() => this.onload?.());
		}
	}
	const canvas = {
		width: 0,
		height: 0,
		getContext: () => ({ drawImage: () => undefined }),
		toBlob(callback: (blob: Blob | null) => void) {
			drawn.size = { width: canvas.width, height: canvas.height };
			callback(new Blob([DOM_PNG]));
		},
	};
	vi.stubGlobal('Image', FakeImage);
	vi.stubGlobal('document', { createElement: () => canvas });
	vi.stubGlobal('URL', {
		createObjectURL: () => 'blob:svg',
		revokeObjectURL: () => undefined,
	});
	return drawn;
}

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('rasterizeSvg', () => {
	it("uses the browser's own canvas when there is a DOM and no rasteriser", async () => {
		const drawn = stubBrowserCanvas();
		const picture = await rasterizeSvg(SVG, 10, 10);
		expect(picture).toStrictEqual({ extension: 'png', bytes: DOM_PNG });
		expect(drawn.size).toStrictEqual({ width: 200, height: 120 });
	});

	it('uses the given rasteriser instead of the DOM, sizing it like PowerPoint', async () => {
		const drawn = stubBrowserCanvas();
		const sizes: Array<{ width: number; height: number }> = [];
		const rasterize = vi.fn<PptxSvgRasterizer>(async (svg, sizeFor) => {
			expect(svg).toBe(SVG);
			sizes.push(sizeFor(100, 60), sizeFor(0, 0), sizeFor(4096, 1024));
			return new Uint8Array([9, 9]);
		});
		const picture = await rasterizeSvg(SVG, 50, 40, rasterize);
		expect(picture).toStrictEqual({ extension: 'png', bytes: new Uint8Array([9, 9]) });
		expect(rasterize).toHaveBeenCalledOnce();
		// Twice the intrinsic size, the fallback size when it declares none, at most 2048 px.
		expect(sizes).toStrictEqual([
			{ width: 200, height: 120 },
			{ width: 100, height: 80 },
			{ width: 2048, height: 512 },
		]);
		expect(drawn.size).toStrictEqual({ width: 0, height: 0 });
	});

	it('returns nothing without a DOM and without a rasteriser', async () => {
		vi.stubGlobal('document', undefined);
		expect(await rasterizeSvg(SVG, 10, 10)).toBeUndefined();
	});

	it('returns nothing when the rasteriser gives up or throws', async () => {
		vi.stubGlobal('document', undefined);
		expect(await rasterizeSvg(SVG, 10, 10, async () => undefined)).toBeUndefined();
		const throwing = () => {
			throw new Error('no canvas');
		};
		expect(await rasterizeSvg(SVG, 10, 10, throwing)).toBeUndefined();
		expect(
			await rasterizeSvg(SVG, 10, 10, async () => Promise.reject(new Error('x'))),
		).toBeUndefined();
	});
});

describe('svgRasterSize', () => {
	it('doubles the intrinsic size and clamps the longest edge to 2048 px', () => {
		expect(svgRasterSize(100, 60, 0, 0)).toStrictEqual({ width: 200, height: 120 });
		expect(svgRasterSize(0, 0, 50, 40)).toStrictEqual({ width: 100, height: 80 });
		expect(svgRasterSize(4096, 1024, 0, 0)).toStrictEqual({ width: 2048, height: 512 });
	});
});
