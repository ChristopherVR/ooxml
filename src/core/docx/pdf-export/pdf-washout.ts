/**
 * Word's picture washout in the PDF export, for body pictures and the watermark painter (which
 * loads on first use, so this stays its own small module). DOM-free.
 */
import { WASHOUT_INTERCEPT, WASHOUT_SLOPE } from '../picture-watermark';
import type { PdfImageResource, PdfVectorPage } from '../../pdf/write/pdf-vector';

const washed = new WeakMap<PdfImageResource, PdfImageResource>();

/** The picture with Word's washout applied to its samples, or undefined for an encoded (JPEG) one. */
export function washedOutResource(resource: PdfImageResource): PdfImageResource | undefined {
	if (resource.filter) return undefined;
	let result = washed.get(resource);
	if (!result) {
		const data = new Uint8Array(resource.data.length);
		for (let index = 0; index < data.length; index++)
			data[index] = Math.min(
				255,
				Math.round(resource.data[index]! * WASHOUT_SLOPE + WASHOUT_INTERCEPT * 255),
			);
		result = { ...resource, data };
		washed.set(resource, result);
	}
	return result;
}

/**
 * Washes out an encoded picture already drawn into the box: a grey layer in Screen mode gives
 * (1 - g) x + g, with g the intercept, so black rises to it and white stays white.
 */
export function paintWashoutLayer(
	page: PdfVectorPage,
	x: number,
	y: number,
	width: number,
	height: number,
): void {
	page.save();
	page.alpha(1, 1, 'Screen');
	page.fillColor(WASHOUT_INTERCEPT, WASHOUT_INTERCEPT, WASHOUT_INTERCEPT);
	page.rect(x, y, width, height);
	page.op('f');
	page.restore();
}
