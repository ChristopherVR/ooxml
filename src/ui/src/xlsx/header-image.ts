// Header and footer pictures for the print preview and Page Layout view: the image bytes of the
// workbook source as a data URL (the print preview is an HTML string, so object URLs would not
// outlive it), sized from the VML shape in points.
import type { HeaderFooterImage, Workbook } from 'ooxml-core/xlsx';

const urls = new WeakMap<Uint8Array, string>();

/** `data:` URL of the image part, or undefined when the source no longer holds it. */
export function headerFooterImageUrl(
	workbook: Workbook,
	image: HeaderFooterImage,
): string | undefined {
	const bytes = workbook.source?.parts.get(image.partName);
	if (!bytes) return undefined;
	const cached = urls.get(bytes);
	if (cached) return cached;
	let binary = '';
	for (let i = 0; i < bytes.length; i += 0x8000)
		binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
	const url = `data:${image.contentType};base64,${btoa(binary)}`;
	urls.set(bytes, url);
	return url;
}

/** CSS pixel size of a header picture at the page scale (96 dpi, so 4/3 pixel per point). */
export function headerFooterImageSize(
	image: HeaderFooterImage,
	scale: number,
): { width: number; height: number } {
	return { width: (image.widthPt * 96 * scale) / 72, height: (image.heightPt * 96 * scale) / 72 };
}
