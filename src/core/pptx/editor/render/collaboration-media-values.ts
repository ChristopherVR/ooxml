/**
 * collaboration-media-values.ts: the values of a slide read from the room, as this viewer
 * shows them.
 *
 * Every media token becomes this page's own URL for the bytes the room holds. A token whose
 * bytes have not arrived, or were too large to send, stays as it is: it draws nothing, and a
 * write of this viewer's slides keeps it in the room instead of deleting it. A `blob:` URL
 * another page wrote is left out. Linked addresses pass through: each viewer gates them where
 * it draws.
 */

import type { CollaborationMedia } from './collaboration-media';
import { isLocalOnlyUrl, mediaTokenId } from './collaboration-media-token';

/** Why a value from the room was left out, when the reader should be told. */
export interface MediaValueIssue {
	readonly field: string;
	readonly reason: 'invalid' | 'withheld';
}

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
		if (resolved.missing === 'invalid') {
			issues.push({ field: key, reason: 'invalid' });
			return undefined;
		}
		if (resolved.missing === 'withheld') {
			issues.push({ field: key, reason: 'withheld' });
		}
		return value;
	}
	if (isLocalOnlyUrl(value) && !media.ownsUrl(value)) {
		issues.push({ field: key, reason: 'invalid' });
		return undefined;
	}
	return value;
}

/** Whether a JSON text from the room holds anything {@link inboundString} would change. */
function needsInbound(json: string): boolean {
	return json.includes('"ooxml-media:') || json.includes('"blob:');
}

/** `JSON.parse(json)` with every string passed through {@link inboundString}. */
export function parseInbound(
	media: CollaborationMedia | undefined,
	json: string,
	issues: MediaValueIssue[],
): unknown {
	if (!media || !needsInbound(json)) {
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
