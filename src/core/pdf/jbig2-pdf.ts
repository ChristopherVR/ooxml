/**
 * The JBIG2Decode filter as the image reader uses it: resolves the /JBIG2Globals stream of the
 * filter's parameters and decodes the image into 1-bit rows (0 for black), or says why not.
 */
import type { PdfDocument } from './document';
import { decodeJbig2, Jbig2Error } from './jbig2';
import { PdfStream, type PdfDict } from './objects';

const MAX_PIXELS = 40_000_000;

export async function jbig2Samples(
	document: PdfDocument,
	data: Uint8Array,
	parms: PdfDict | undefined,
	width: number,
	height: number,
): Promise<{ bytes: Uint8Array; fax: true } | string> {
	let globals: Uint8Array | undefined;
	const value = await document.resolve(parms?.get('JBIG2Globals'));
	if (value instanceof PdfStream) {
		const decoded = await document.streamData(value);
		if (decoded.imageFilter) return 'its JBIG2 globals use an unsupported filter';
		globals = decoded.bytes;
	}
	try {
		const bytes = decodeJbig2(data, globals, width, height, { maxPixels: MAX_PIXELS });
		return { bytes, fax: true };
	} catch (error) {
		if (!(error instanceof Jbig2Error)) throw error;
		const detail = error.message.replace(/\.$/, '');
		return `its JBIG2 data could not be decoded (${detail.charAt(0).toLowerCase()}${detail.slice(1)})`;
	}
}
