/**
 * Which URLs a document resource field (a picture, a media file, an embedded file) may hold
 * when the value comes from outside the loader, and which `data:` URLs a viewer may open in
 * a browser tab. Shared by every format; the companion of `safe-href.ts`, which covers
 * hyperlinks.
 *
 * The type of a `data:` URL is always read from the URL itself, never from a separate MIME
 * field, because the browser decides how to render the URL from its own header.
 */
import { hyperlinkPolicy } from './safe-href';

/** What a resource field holds, which decides the `data:` types it accepts. */
export type ResourceKind = 'image' | 'media' | 'model' | 'file';

/** Where a resource field's value may point besides a `data:` URL. */
export interface ResourceRule {
	readonly kind: ResourceKind;
	/** An `http:` or `https:` URL (a linked external picture or media file). */
	readonly web?: boolean;
	/** A path without a scheme (a part inside the package). */
	readonly path?: boolean;
}

/** The header of a `data:` URL. `mime` is lower case and has no parameters. */
export interface DataUrlHeader {
	readonly mime: string;
	readonly base64: boolean;
}

const TOKEN = "[a-z0-9!#$&^_.+'-]+";
const MIME = new RegExp(`^${TOKEN}/${TOKEN}$`, 'u');
// `name=value`, or a bare `name` such as the common `;utf8`.
const PARAMETER = new RegExp(`^${TOKEN}(?:=(?:${TOKEN}|"[^"\\\\]*"))?$`, 'u');
/** Longer headers are not produced by any writer and are not parsed. */
const MAX_HEADER = 256;

/**
 * A well-formed MIME type (`type/subtype`, at most 255 characters, no parameters), lower
 * case, or `undefined`.
 */
export function normalizeMimeType(value: unknown): string | undefined {
	if (typeof value !== 'string' || value.length > 255) return undefined;
	const mime = value.trim().toLowerCase();
	return MIME.test(mime) ? mime : undefined;
}

/** Whether `value` is a well-formed content type: a MIME type with optional parameters. */
export function isWellFormedContentType(value: unknown): value is string {
	if (typeof value !== 'string' || value.length > 255) return false;
	const [type = '', ...parameters] = value.toLowerCase().split(';');
	return (
		normalizeMimeType(type) !== undefined &&
		parameters.every((parameter) => PARAMETER.test(parameter.trim()))
	);
}

/**
 * Parse the header of a `data:` URL without reading its payload. Returns `undefined` for
 * anything that is not a well-formed `data:` URL. A URL with no type is `text/plain`, as in
 * a browser.
 */
export function parseDataUrlHeader(url: unknown): DataUrlHeader | undefined {
	if (typeof url !== 'string' || url.slice(0, 5).toLowerCase() !== 'data:') return undefined;
	const comma = url.indexOf(',', 5);
	if (comma < 0 || comma > MAX_HEADER) return undefined;
	const parts = url.slice(5, comma).toLowerCase().split(';');
	const base64 = parts.length > 1 && parts.at(-1) === 'base64';
	if (base64) parts.pop();
	const [type = '', ...parameters] = parts;
	const mime = type === '' ? 'text/plain' : normalizeMimeType(type);
	if (!mime || type !== type.trim()) return undefined;
	if (!parameters.every((parameter) => PARAMETER.test(parameter))) return undefined;
	return { mime, base64 };
}

/**
 * Types a viewer may open in a new browser tab from its own origin. Each one is shown by
 * the browser as an inert document: no markup, no script. Everything else is download only.
 */
const BROWSER_OPENABLE = new Set([
	'application/pdf',
	'application/json',
	'text/plain',
	'text/csv',
	'image/png',
	'image/jpeg',
	'image/gif',
	'image/webp',
	'image/bmp',
	'image/avif',
]);

/** Whether a payload of this type may be opened in a browser tab (see {@link BROWSER_OPENABLE}). */
export function isBrowserOpenableType(mime: string | undefined): boolean {
	const normalized = normalizeMimeType(mime?.split(';', 1)[0]);
	return normalized !== undefined && BROWSER_OPENABLE.has(normalized);
}

/** Whether `url` is a `data:` URL whose own type may be opened in a browser tab. */
export function isBrowserOpenableDataUrl(url: unknown): boolean {
	const header = parseDataUrlHeader(url);
	return header !== undefined && BROWSER_OPENABLE.has(header.mime);
}

const isImage = (mime: string): boolean => mime.startsWith('image/');
const isOpaque = (mime: string): boolean => mime === 'application/octet-stream';
// Loaders label a media or model part with an unknown extension as an image, so those
// kinds accept image types too.
const KIND_ACCEPTS: Readonly<Record<ResourceKind, (mime: string) => boolean>> = {
	image: isImage,
	media: (mime) =>
		mime.startsWith('audio/') || mime.startsWith('video/') || isImage(mime) || isOpaque(mime),
	model: (mime) => mime.startsWith('model/') || isImage(mime) || isOpaque(mime),
	file: () => true,
};

/**
 * Whether `value` may be kept in a resource field described by `rule`. A `data:` URL must be
 * well formed and of a type the field's kind holds; `http(s)` URLs and package paths only
 * where the rule allows them. Script schemes, other schemes and disguised schemes never pass.
 */
export function isAllowedResourceUrl(value: unknown, rule: ResourceRule): boolean {
	if (typeof value !== 'string' || value.length === 0) return false;
	const header = parseDataUrlHeader(value);
	if (header) return KIND_ACCEPTS[rule.kind](header.mime);
	const policy = hyperlinkPolicy(value);
	if (!policy.store) return false;
	const scheme = /^([a-z][a-z0-9+.-]*):/iu.exec(value)?.[1]?.toLowerCase();
	if (scheme === 'http' || scheme === 'https') return rule.web === true;
	return scheme === undefined && rule.path === true;
}
