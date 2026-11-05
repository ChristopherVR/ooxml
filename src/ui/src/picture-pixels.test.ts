// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { cachedPicturePixels, probePicturePixels } from './picture-pixels';

/**
 * Processed pictures stay cached even when their data URL is far larger than the shared render
 * cache's 8 MB cap: an uncached result was recomputed on every render. The canvas is a stub; only
 * the caching and the number of PNG encodes are checked.
 */
const tint = [{ kind: 'standard', name: 'tint', hue: 0, amount: 50 }] as never;
const big = `data:image/png;base64,${'A'.repeat(6 * 1024 * 1024)}`;

afterEach(() => vi.restoreAllMocks());

describe('probePicturePixels', () => {
	it('keeps a multi-megabyte result cached and encodes it once', async () => {
		const toDataURL = vi.fn(() => big);
		const context = {
			drawImage: vi.fn(),
			clearRect: vi.fn(),
			getImageData: vi.fn(() => ({ data: new Uint8ClampedArray(4) })),
			putImageData: vi.fn(),
		};
		const getContext = vi.fn(() => context);
		const realCreate = document.createElement.bind(document);
		vi.spyOn(document, 'createElement').mockImplementation(((tag: string) => {
			if (tag === 'canvas')
				return { width: 0, height: 0, getContext, toDataURL } as unknown as HTMLCanvasElement;
			if (tag === 'img') {
				const image = { naturalWidth: 2, naturalHeight: 2 } as unknown as HTMLImageElement & {
					onload?: () => void;
				};
				Object.defineProperty(image, 'src', {
					set: () => {
						queueMicrotask(() => image.onload?.());
					},
				});
				return image;
			}
			return realCreate(tag);
		}) as typeof document.createElement);

		const result = await probePicturePixels(document, 'blob:picture-1', tint);
		expect(result?.url).toBe(big);
		expect(toDataURL).toHaveBeenCalledOnce();
		expect(getContext).toHaveBeenCalledWith('2d', { willReadFrequently: true });
		expect(cachedPicturePixels('blob:picture-1', tint)?.url).toBe(big);
	});
});
