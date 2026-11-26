import { encodeBase64 } from '../../digest/base64';
import { decodePng } from '../../drawingml/raster-decode';
import { encodePng } from '../../drawingml/raster-encode';
import { MANIFEST_BYTES } from './validation';

export function pngPreview(bytes: Uint8Array): string | undefined {
	const bitmap = decodePng(bytes, {
		maxImageBytes: MANIFEST_BYTES,
		maxDimension: 512,
		maxPixels: 512 * 512,
	});
	if (!bitmap) return;
	const { width, height, rgb, alpha } = bitmap;
	// Emit only bounded pixel data. The browser never parses the package's PNG chunks.
	const preview = encodePng(width, height, rgb, alpha);
	if (preview.byteLength <= MANIFEST_BYTES) return `data:image/png;base64,${encodeBase64(preview)}`;
}
