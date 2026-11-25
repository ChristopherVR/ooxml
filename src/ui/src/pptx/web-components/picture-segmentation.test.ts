import { describe, expect, it, vi } from 'vitest';
import { segmentPicture } from './picture-segmentation';

describe('picture segmentation', () => {
	it('removes the connected background and honors foreground and background marks', () => {
		function sample(width: number, height: number): ImageData {
			return { width, height, colorSpace: 'srgb', data: new Uint8ClampedArray(width * height * 4) };
		}
		vi.stubGlobal('ImageData', sample);
		try {
			const image = sample(20, 20);
			for (let y = 0; y < 20; y++) {
				for (let x = 0; x < 20; x++) {
					const offset = (y * 20 + x) * 4;
					image.data[offset] = x >= 5 && x < 15 && y >= 5 && y < 15 ? 230 : 15;
					image.data[offset + 3] = 255;
				}
			}
			const alphaAt = (mask: ImageData, x: number, y: number) => mask.data[(y * 20 + x) * 4 + 3];
			const guess = segmentPicture(image, { foreground: [], background: [] });
			expect(alphaAt(guess.mask, 1, 1)).toBe(0);
			expect(alphaAt(guess.mask, 10, 10)).toBe(255);
			expect(guess.removal).toMatchObject({ top: 0.25, left: 0.25, bottom: 0.75, right: 0.75 });
			const result = segmentPicture(image, {
				foreground: [{ x1: 0.05, y1: 0.05, x2: 0.05, y2: 0.05 }],
				background: [{ x1: 0.5, y1: 0.5, x2: 0.5, y2: 0.5 }],
			});
			expect(alphaAt(result.mask, 1, 1)).toBe(255);
			expect(alphaAt(result.mask, 10, 10)).toBe(0);
		} finally {
			vi.unstubAllGlobals();
		}
	});
});
