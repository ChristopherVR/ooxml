import { inspectBmp } from '../drawingml/bmp-inspection';
import {
	checkedImageLimits,
	inspectRasterImageStructure,
	RasterImageError,
	type ImageLimits,
	type RasterImageInfo,
} from '../drawingml/raster-image';
import { inspectSvgPicture } from '../drawingml/static-svg-picture';
import { hasTiffSignature, inspectTiff } from '../drawingml/tiff-picture';
import { inspectVisioEmfAdmission } from './emf-admission';
import { emfInputView } from './emf-admission-input';

import {
	VISIO_EMF_CONVERSION_LIMITS,
	VISIO_EMF_CONVERSION_MAX_DIMENSION,
} from './emf-admission-types';

export type VisioPictureMime =
	| RasterImageInfo['mimeType']
	| 'image/bmp'
	| 'image/tiff'
	| 'image/svg+xml'
	| 'image/x-emf';
export type VisioPictureInfo = Omit<RasterImageInfo, 'mimeType'> & { mimeType: VisioPictureMime };

export function inspectPicture(bytes: Uint8Array, limits: ImageLimits, allowEmf: boolean) {
	const view = emfInputView(bytes);
	if (!view) throw new RasterImageError('invalid-image', 'Expected fixed, unshared picture bytes.');
	const data = new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
	if (data[0] === 0x42 && data[1] === 0x4d) {
		const [pixelWidth, pixelHeight] = inspectBmp(view, limits);
		return { mimeType: 'image/bmp' as const, pixelWidth, pixelHeight };
	}
	if (hasTiffSignature(view)) {
		const tiff = inspectTiff(view, limits);
		return {
			mimeType: 'image/tiff' as const,
			pixelWidth: tiff.pixelWidth,
			pixelHeight: tiff.pixelHeight,
			tiff,
		};
	}
	if (allowEmf && view.byteLength >= 4 && view.getUint32(0, true) === 1) {
		const result = inspectVisioEmfAdmission(data, VISIO_EMF_CONVERSION_LIMITS);
		const { header } = result;
		if (result.status !== 'admitted' || !header) {
			const code =
				result.status === 'unsupported'
					? 'unsupported-image'
					: result.status === 'budget-exceeded'
						? 'image-limit'
						: 'invalid-image';
			throw new RasterImageError(code, result.diagnostics[0]?.message ?? 'Invalid EMF picture.');
		}
		const { pixelWidth, pixelHeight } = header;
		if (Math.max(pixelWidth, pixelHeight) > VISIO_EMF_CONVERSION_MAX_DIMENSION)
			throw new RasterImageError('image-limit', 'EMF dimensions exceed 2048 pixels.');
		return { mimeType: 'image/x-emf' as const, pixelWidth, pixelHeight };
	}
	const prefix = new TextDecoder().decode(data.subarray(0, 64)).trimStart();
	if (prefix.startsWith('<')) {
		const { pixelWidth, pixelHeight, root } = inspectSvgPicture(data, limits);
		return { mimeType: 'image/svg+xml' as const, pixelWidth, pixelHeight, svg: root };
	}
	return inspectRasterImageStructure(data, limits);
}

export function inspectPictureInput(bytes: Uint8Array) {
	return inspectPicture(bytes, checkedImageLimits(bytes, {}), true);
}
