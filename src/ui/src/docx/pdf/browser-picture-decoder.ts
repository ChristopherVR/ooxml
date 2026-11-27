/**
 * Decodes the picture formats the PDF export does not read itself (GIF, BMP, TIFF, WebP) with the
 * browser's own image decoder, into RGBA. Nothing is fetched: the bytes are decoded in memory.
 * Returns `undefined` where the browser cannot decode the format (EMF, WMF) or has no canvas.
 */
import type { DecodedPicture } from 'ooxml-core/docx/pdf-export';

/** The longest side a decoded picture keeps, as the other PDF exports cap it. */
const MAX_SIDE = 4096;

export async function decodePictureInBrowser(
	bytes: Uint8Array,
	contentType: string,
): Promise<DecodedPicture | undefined> {
	if (typeof createImageBitmap !== 'function' || typeof document === 'undefined') return undefined;
	let bitmap: ImageBitmap;
	try {
		bitmap = await createImageBitmap(new Blob([new Uint8Array(bytes)], { type: contentType }));
	} catch {
		return undefined;
	}
	try {
		const scale = Math.min(1, MAX_SIDE / Math.max(bitmap.width, bitmap.height));
		const width = Math.max(1, Math.round(bitmap.width * scale));
		const height = Math.max(1, Math.round(bitmap.height * scale));
		const canvas = document.createElement('canvas');
		canvas.width = width;
		canvas.height = height;
		const context = canvas.getContext('2d');
		if (!context) return undefined;
		context.drawImage(bitmap, 0, 0, width, height);
		return { width, height, rgba: context.getImageData(0, 0, width, height).data };
	} finally {
		bitmap.close();
	}
}
