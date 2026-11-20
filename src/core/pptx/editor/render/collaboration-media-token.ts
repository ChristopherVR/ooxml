/**
 * collaboration-media-token.ts: how a page-local URL travels through a collaboration room.
 *
 * A `blob:` URL resolves only inside the page that made it. The loader gives every picture,
 * poster frame and picture fill such a URL, so writing it into the shared document would hand
 * every other participant a source that never loads. Instead each one is written as a media
 * token, `ooxml-media:<id>`, and its bytes are put in the room once under that id (see
 * collaboration-media.ts); every peer turns the bytes into a URL of its own.
 *
 * Object URLs belong to the page, not to one document, so the mapping between a URL and its
 * token is kept per page here. A URL this page minted from a room's bytes maps back to the
 * token it came from, so writing it again never uploads the bytes a second time.
 */

const TOKEN_PREFIX = 'ooxml-media:';
const TOKEN = /^ooxml-media:([0-9a-f]{32})$/u;

const idByUrl = new Map<string, string>();
const urlById = new Map<string, string>();
let collecting: Set<string> | undefined;

/** Whether `value` is a page-local URL that must not enter the shared document as is. */
export function isLocalOnlyUrl(value: unknown): value is string {
	return typeof value === 'string' && value.startsWith('blob:');
}

/** The id inside a well-formed media token, or `undefined`. */
export function mediaTokenId(value: unknown): string | undefined {
	return typeof value === 'string' && value.startsWith(TOKEN_PREFIX)
		? TOKEN.exec(value)?.[1]
		: undefined;
}

export function mediaToken(id: string): string {
	return `${TOKEN_PREFIX}${id}`;
}

function newId(): string {
	const bytes = new Uint8Array(16);
	globalThis.crypto.getRandomValues(bytes);
	return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/** Record that `url` (a URL this page made) stands for media `id`; the first URL of an id stays. */
export function rememberMediaUrl(url: string, id: string): void {
	idByUrl.set(url, id);
	if (!urlById.has(id)) {
		urlById.set(id, url);
	}
}

/** Forget `url`, for example once it has been revoked. */
export function forgetMediaUrl(url: string): void {
	const id = idByUrl.get(url);
	idByUrl.delete(url);
	if (id !== undefined && urlById.get(id) === url) {
		urlById.delete(id);
	}
}

/** The media id a URL of this page stands for, if it was shared or minted. */
export function mediaIdOf(url: string): string | undefined {
	return idByUrl.get(url);
}

/** The URL this page holds for media `id`, if it made or minted one. */
export function mediaUrlOf(id: string): string | undefined {
	return urlById.get(id);
}

/**
 * The token a page-local URL is written as. The same URL always gets the same token, and a
 * URL minted from a token gets that token back.
 */
export function shareMediaUrl(url: string): string {
	let id = idByUrl.get(url);
	if (id === undefined) {
		id = newId();
		rememberMediaUrl(url, id);
	}
	collecting?.add(id);
	return mediaToken(id);
}

/** `value` as it is written to the shared document: a page-local URL becomes its token. */
export function sharedValue<T>(value: T): T | string {
	return isLocalOnlyUrl(value) ? shareMediaUrl(value) : value;
}

const sharedReplacer = (_key: string, value: unknown): unknown => sharedValue(value);

/** `JSON.stringify(value)`, with every page-local URL inside written as its token. */
export function stringifyShared(value: unknown): string {
	const json = JSON.stringify(value);
	return json !== undefined && json.includes('"blob:')
		? JSON.stringify(value, sharedReplacer)
		: json;
}

/**
 * Run `write` and return the ids of the media tokens it wrote, so the caller can put their
 * bytes in the room. Writes are synchronous, so the collector sees exactly this call's tokens.
 */
export function collectSharedMedia(write: () => void): Set<string> {
	const outer = collecting;
	const ids = new Set<string>();
	collecting = ids;
	try {
		write();
	} finally {
		collecting = outer;
		if (outer) {
			for (const id of ids) {
				outer.add(id);
			}
		}
	}
	return ids;
}
