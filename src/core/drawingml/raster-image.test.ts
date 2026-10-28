import { expect, it } from 'vitest';
import { inspectRasterImage, RasterImageError } from './raster-image';
import { inspectVisioRasterImage, VisioImageError } from '../visio/media';
import { pictureGridPng } from '../diagram/test-support/picture-grid';

it('shares intrinsic size admission while retaining the Visio public error contract', () => {
	const bytes = pictureGridPng();
	expect(inspectRasterImage(bytes)).toEqual({
		mimeType: 'image/png',
		pixelWidth: 48,
		pixelHeight: 32,
	});
	expect(inspectVisioRasterImage(bytes)).toEqual(inspectRasterImage(bytes));
	const corrupt = bytes.slice();
	corrupt[30] = (corrupt[30] ?? 0) ^ 1;
	expect(() => inspectRasterImage(corrupt)).toThrow(RasterImageError);
	expect(() => inspectVisioRasterImage(corrupt)).toThrow(VisioImageError);
	try {
		inspectVisioRasterImage(corrupt);
	} catch (error) {
		expect((error as Error).name).toBe('VisioImageError');
	}
});
