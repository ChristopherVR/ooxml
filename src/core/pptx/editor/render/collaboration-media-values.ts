/**
 * collaboration-media-values.ts: the values of a slide read from the room, as this viewer
 * shows them.
 *
 * Every media token becomes this page's own URL for the bytes the room holds (or is left
 * out while they have not arrived), a `blob:` URL another page wrote is left out, and a
 * linked picture or media file is left out when this viewer's remote content setting does
 * not allow it, whatever the participant who shared it allows.
 */

import type { CollaborationMedia } from './collaboration-media';
import { isLocalOnlyUrl, mediaTokenId } from './collaboration-media-token';

/** Why a value from the room was left out, when the reader should be told. */
export interface MediaValueIssue {
	readonly field: string;
	readonly reason: 'invalid' | 'withheld';
}

/**
 * Keys that hold a picture, media file or embedded file a viewer fetches or shows, wherever
 * they appear in a slide (element fields, fills, text runs, bullets, table cells).
 */
const RESOURCE_KEYS: ReadonlySet<string> = new Set([
	'imageData',
	'svgData',
	'posterImage',
	'posterFrameData',
	'previewImage',
	'previewImageData',
	'mediaData',
	'modelData',
	'oleEmbeddedData',
	'backgroundImage',
	'fillImageUrl',
	'textFillBlipUrl',
	'backgroundImageFillData',
	'imageDataUrl',
	'imageUrl',
	'soundData',
]);

const WEB_URL = /^https?:/iu;

/** One string from the room as this viewer uses it, or `undefined` to leave it out. */
export function inboundString(
	media: CollaborationMedia,
	key: string,
	value: string,
	issues: MediaValueIssue[],
): string | undefined {
	if (value.startsWith('ooxml-media:')) {
		const id = mediaTokenId(value);
		const resolved = id === undefined ? { missing: 'invalid' as const } : media.resolve(id);
		if ('url' in resolved) {
			return resolved.url;
		}
		if (resolved.missing !== 'pending') {
			issues.push({ field: key, reason: resolved.missing });
		}
		return undefined;
	}
	if (isLocalOnlyUrl(value) && !media.ownsUrl(value)) {
		issues.push({ field: key, reason: 'invalid' });
		return undefined;
	}
	if (!media.allowExternalContent && RESOURCE_KEYS.has(key) && WEB_URL.test(value)) {
		return undefined;
	}
	return value;
}

/** Whether a JSON text from the room holds anything {@link inboundString} would change. */
function needsInbound(media: CollaborationMedia, json: string): boolean {
	return (
		json.includes('"ooxml-media:') ||
		json.includes('"blob:') ||
		(!media.allowExternalContent && /"https?:/iu.test(json))
	);
}

/** `JSON.parse(json)` with every string passed through {@link inboundString}. */
export function parseInbound(
	media: CollaborationMedia | undefined,
	json: string,
	issues: MediaValueIssue[],
): unknown {
	if (!media || !needsInbound(media, json)) {
		return JSON.parse(json);
	}
	return JSON.parse(json, (key, value: unknown) =>
		typeof value === 'string' ? inboundString(media, key, value, issues) : value,
	);
}

/** A copy of `value` with every string passed through {@link inboundString}. */
export function mapInbound(
	media: CollaborationMedia,
	value: unknown,
	issues: MediaValueIssue[],
	key = '',
): unknown {
	if (typeof value === 'string') {
		return inboundString(media, key, value, issues);
	}
	if (Array.isArray(value)) {
		return value.map((item) => mapInbound(media, item, issues, key));
	}
	if (typeof value !== 'object' || value === null) {
		return value;
	}
	const result: Record<string, unknown> = {};
	for (const [field, item] of Object.entries(value)) {
		const mapped = mapInbound(media, item, issues, field);
		if (mapped !== undefined) {
			result[field] = mapped;
		}
	}
	return result;
}
