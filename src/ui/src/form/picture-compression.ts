export interface PictureCompression {
	width: number;
	height: number;
	ppi: number;
	crop: readonly [left: number, top: number, right: number, bottom: number];
	deleteCrop: boolean;
}

interface OpenPicture {
	bitmap: ImageBitmap;
	source: Blob;
}

export async function openRaster(
	source: Blob,
	header: { width: number; height: number },
): Promise<OpenPicture> {
	const bitmap = await createImageBitmap(source);
	if (
		!(bitmap.width === header.width && bitmap.height === header.height) &&
		!(bitmap.width === header.height && bitmap.height === header.width)
	) {
		bitmap.close();
		throw new Error('Decoded raster dimensions differ from the header');
	}
	return { bitmap, source };
}

export function contextFor(width: number, height: number): OffscreenCanvasRenderingContext2D {
	const context = new OffscreenCanvas(width, height).getContext('2d');
	if (!context) throw new Error('Offscreen canvas is unavailable');
	return context;
}

/** Crop in source pixels and downsample to the target resolution without upscaling. */
export async function compressRaster(
	{ bitmap, source }: OpenPicture,
	options: PictureCompression,
): Promise<{ kind: 'compressed'; image: Blob; deletedCrop: boolean }> {
	const [left, top, right, bottom] = options.crop;
	const visibleX = Math.max(0.01, 1 - left - right);
	const visibleY = Math.max(0.01, 1 - top - bottom);
	const deletedCrop =
		options.deleteCrop &&
		options.crop.every((edge) => edge >= 0) &&
		left + right < 1 &&
		top + bottom < 1 &&
		options.crop.some((edge) => edge > 0);
	const x = deletedCrop ? left * bitmap.width : 0;
	const y = deletedCrop ? top * bitmap.height : 0;
	const width = deletedCrop ? (1 - left - right) * bitmap.width : bitmap.width;
	const height = deletedCrop ? (1 - top - bottom) * bitmap.height : bitmap.height;
	const xScale = (options.width * options.ppi) / (96 * bitmap.width * visibleX);
	const yScale = (options.height * options.ppi) / (96 * bitmap.height * visibleY);
	if (!Number.isFinite(xScale) || !Number.isFinite(yScale))
		throw new Error('Invalid picture frame');
	const scale = Math.min(1, Math.max(0, xScale, yScale));
	const outputWidth = Math.max(1, Math.round(width * scale));
	const outputHeight = Math.max(1, Math.round(height * scale));
	const context = contextFor(outputWidth, outputHeight);
	context.imageSmoothingQuality = 'high';
	context.drawImage(bitmap, x, y, width, height, 0, 0, outputWidth, outputHeight);
	const image = await context.canvas.convertToBlob({ type: source.type, quality: 0.84 });
	return { kind: 'compressed', image, deletedCrop };
}
