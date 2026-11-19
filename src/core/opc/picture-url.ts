/**
 * The URL a viewer shows a picture's bytes from. A `blob:` URL has the origin of the page that made
 * it, and a browser's "Open image in new tab" renders it as a document of that origin; for an SVG
 * picture, a scriptable document the file controls, that would run in the application's origin.
 * So only a type the browser shows inertly (a raster image) gets a `blob:` URL; any other type,
 * SVG above all, gets a `data:` URL, whose origin is opaque when opened on its own. Both render
 * the same bytes in an `<img>`, so nothing is lost but the memory of the base64 text.
 */
import { encodeBase64 } from '../digest/base64';
import { isBrowserOpenableType, normalizeMimeType } from './safe-resource-url';

/** Makes a `blob:` URL; the default is the global `URL.createObjectURL`, when there is one. */
export type ObjectUrlFactory = (blob: Blob) => string;

const globalObjectUrl = (): ObjectUrlFactory | null => {
	const url = globalThis.URL as { createObjectURL?: (blob: Blob) => string } | undefined;
	return typeof url?.createObjectURL === 'function' && typeof globalThis.Blob === 'function'
		? (blob) => url.createObjectURL!(blob)
		: null;
};

/** Whether a picture of this type may be shown from a `blob:` URL of the page's own origin. */
export function mayUseOriginBlobUrl(mime: string): boolean {
	return isBrowserOpenableType(mime) && normalizeMimeType(mime)!.startsWith('image/');
}

/**
 * A URL for `bytes`, a picture of type `mime`: a `blob:` URL from `objectUrl` for an inert raster
 * type, otherwise (or when `objectUrl` is null, as it is without a global one) a `data:` URL.
 * Revoke it as any object URL; revoking a `data:` URL does nothing.
 */
export function pictureUrl(
	bytes: Uint8Array,
	mime: string,
	objectUrl: ObjectUrlFactory | null = globalObjectUrl(),
): string {
	const type = normalizeMimeType(mime) ?? 'application/octet-stream';
	if (objectUrl && mayUseOriginBlobUrl(type)) return objectUrl(new Blob([bytes.slice()], { type }));
	return `data:${type};base64,${encodeBase64(bytes)}`;
}
