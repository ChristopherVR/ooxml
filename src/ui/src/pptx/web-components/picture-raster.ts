import { contextFor } from '../../form/picture-compression';
export { openRaster, compressRaster } from '../../form/picture-compression';
import { segmentPicture, type PictureMarks } from './picture-segmentation';
import type { PictureWorkerResult } from './picture-worker-protocol';

function scaledSize(bitmap: ImageBitmap, maxEdge: number): [number, number] {
	const scale = Math.min(1, maxEdge / Math.max(bitmap.width, bitmap.height));
	return [
		Math.max(1, Math.round(bitmap.width * scale)),
		Math.max(1, Math.round(bitmap.height * scale)),
	];
}

export function rasterSegmentationSample(bitmap: ImageBitmap): ImageData {
	const [width, height] = scaledSize(bitmap, 256);
	const sample = contextFor(width, height);
	sample.drawImage(bitmap, 0, 0, width, height);
	return sample.getImageData(0, 0, width, height);
}

export async function removeRasterBackground(
	bitmap: ImageBitmap,
	sample: ImageData,
	marks: PictureMarks,
	preview: boolean,
): Promise<Extract<PictureWorkerResult, { kind: 'removed' }>> {
	const { mask, removal } = segmentPicture(sample, marks);
	const alpha = contextFor(mask.width, mask.height);
	alpha.putImageData(mask, 0, 0);
	const [width, height] = scaledSize(bitmap, preview ? 550 : Infinity);
	const output = contextFor(width, height);
	output.drawImage(bitmap, 0, 0, width, height);
	output.globalCompositeOperation = 'destination-in';
	output.drawImage(alpha.canvas, 0, 0, width, height);
	const image = await output.canvas.convertToBlob();
	return { kind: 'removed', image, removal };
}
