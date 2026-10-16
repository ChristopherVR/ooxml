import type { PptxImageProperties } from '../core/types/image';

interface ImageCropFrame extends PptxImageProperties {
	width: number;
	height: number;
}

/** Map the source crop into a clipped viewport without changing the picture's frame. */
export function wrapSvgImageCrop(frame: ImageCropFrame, image: string): string {
	const { width, height, cropLeft = 0, cropTop = 0, cropRight = 0, cropBottom = 0 } = frame;
	if (cropLeft === 0 && cropTop === 0 && cropRight === 0 && cropBottom === 0) {
		return image;
	}
	const x = width * cropLeft;
	const y = height * cropTop;
	const visibleWidth = width * (1 - cropLeft - cropRight);
	const visibleHeight = height * (1 - cropTop - cropBottom);
	if (
		![width, height, x, y, visibleWidth, visibleHeight].every(Number.isFinite) ||
		width <= 0 ||
		height <= 0 ||
		visibleWidth <= 0 ||
		visibleHeight <= 0
	) {
		return image;
	}
	// Signed insets also represent padding, so do not clamp them to [0, 1].
	return (
		`<svg width="${width}" height="${height}" viewBox="${x} ${y} ${visibleWidth} ${visibleHeight}"` +
		` preserveAspectRatio="none" overflow="hidden">${image}</svg>`
	);
}
