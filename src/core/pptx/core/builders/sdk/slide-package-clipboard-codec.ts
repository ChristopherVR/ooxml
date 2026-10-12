import { base64ToBytes, bytesToBase64 } from '../../../converter/json/json-binary-codec';

export const SLIDE_PACKAGE_CLIPBOARD_MIME = 'web application/vnd.ooxml.pptx-slide-package';
const MARKER = 'ooxml.pptx.slide-package.v1\n';
/** Portable text transport for a selected-slides-only OOXML package. */
export function serializeSlidePackageClipboard(bytes: Uint8Array): string {
	return MARKER + bytesToBase64(bytes);
}
/** Foreign clipboard text returns null; malformed recognized payloads throw. */
export function deserializeSlidePackageClipboard(text: string): Uint8Array | null {
	if (!text.startsWith(MARKER)) return null;
	const payload = text.slice(MARKER.length);
	if (
		!payload ||
		payload.length % 4 ||
		!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/u.test(payload)
	)
		throw new Error('Copied slide package is malformed');
	const bytes = base64ToBytes(payload);
	if (bytes[0] !== 0x50 || bytes[1] !== 0x4b)
		throw new Error('Copied slides do not contain an OOXML archive');
	return bytes;
}
