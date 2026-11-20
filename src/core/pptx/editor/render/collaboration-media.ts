/**
 * collaboration-media.ts: the deck's pictures and media inside a collaboration room.
 *
 * The slides in the room name their pictures and media by media token (see
 * collaboration-media-token.ts); the bytes behind each token live once in the `pptx:media`
 * map, `{ t: <MIME type>, b: <bytes> }` under the token's id. The participant whose page
 * holds the URL puts the bytes there after writing the slides that use it; every other
 * participant turns them into a URL of its own when it reads those slides.
 *
 * Limits. The whole room reaches a joiner in its first sync, and every client refuses an
 * update over `MAX_UPDATE_BYTES` (`ooxml-core/collab`), so the media of one room are kept
 * under {@link MAX_SHARED_MEDIA_BYTES}. A file that would go over it is not sent: its entry
 * keeps only its type and size (`{ t, n }`), and the participants who read it are told it
 * was too large to share. Entries no slide refers to any more are removed before new bytes
 * are added.
 *
 * Linked content (a picture or video at an `http(s)` URL) is never fetched and re-sent: it
 * stays a link, and each participant shows it only when its own remote content setting
 * allows it (`allowExternalContent`), as the loader does for a deck opened from a file.
 */

import { encodeBase64 } from '../../../digest/base64';
import { mayUseOriginBlobUrl } from '../../../opc/picture-url';
import {
	normalizeMimeType,
	resourceKindAccepts,
	type ResourceKind,
} from '../../../opc/safe-resource-url';
import {
	forgetMediaUrl,
	isLocalOnlyUrl,
	mediaIdOf,
	mediaUrlOf,
	rememberMediaUrl,
} from './collaboration-media-token';
import { referencedMediaIds } from './collaboration-media-scan';
import type { YDeepObserver, YDocLike, YMapLike, YTransactionLike } from './collaboration-sync';

export const YDOC_MEDIA_KEY = 'pptx:media';

/** Half of the largest update a client accepts, leaving the other half for the slides. */
export const MAX_SHARED_MEDIA_BYTES = 16 * 1024 * 1024;

/**
 * The budget of a peer-to-peer (y-webrtc) room. y-webrtc sends a joiner the whole room in one
 * data channel message, and Chromium refuses a message over 256 KiB, so media get half of it.
 */
export const MAX_PEER_TO_PEER_MEDIA_BYTES = 128 * 1024;

export interface CollaborationMediaOptions {
	/**
	 * The playable URL this page resolved for an embedded audio or video part (the binding's
	 * `mediaDataUrls`), so a media element that names only its part can share its bytes.
	 */
	localMediaUrl?: ((path: string) => string | undefined) | undefined;
	/**
	 * This viewer's remote content setting; linked pictures and media are left out when false
	 * (the default, as in the core loader). Like the loader's, it applies to what is read next.
	 */
	allowExternalContent?: boolean | undefined;
	/** The room's media budget; {@link MAX_SHARED_MEDIA_BYTES} unless the transport needs less. */
	maxBytes?: number | undefined;
}

/** What a media token stands for here. */
export type MediaResolution =
	| { readonly url: string }
	| { readonly missing: 'pending' | 'withheld' | 'invalid' };

export interface CollaborationMedia {
	/** Set the given options; options left out keep their value. */
	configure: (options: CollaborationMediaOptions) => void;
	readonly allowExternalContent: boolean;
	/** The page-local URL that plays the media part `path`, when this page has one. */
	localMediaUrl: (path: string) => string | undefined;
	resolve: (id: string) => MediaResolution;
	/** Whether `url` is one this page made for a picture or media file a field of `kind` holds. */
	accepts: (url: unknown, kind: ResourceKind) => boolean;
	/** Whether `url` is a page-local URL this page made or minted from the room. */
	ownsUrl: (url: string) => boolean;
	/** Put the bytes of media `ids` in the room, in transactions tagged `origin`. */
	upload: (ids: Iterable<string>, origin: unknown) => Promise<void>;
	/** Call `onChange` when the room's media change. */
	observe: (onChange: YDeepObserver) => () => void;
}

interface MediaEntry {
	readonly t: string;
	readonly b?: Uint8Array;
	readonly n?: number;
}

interface ObservableMap extends YMapLike {
	observe?: (handler: (event: unknown, transaction: YTransactionLike) => void) => void;
	unobserve?: (handler: (event: unknown, transaction: YTransactionLike) => void) => void;
}

const readEntry = (value: unknown): MediaEntry | 'invalid' | undefined => {
	if (value === undefined) {
		return undefined;
	}
	if (typeof value !== 'object' || value === null) {
		return 'invalid';
	}
	const { t, b, n } = value as Record<string, unknown>;
	const type = normalizeMimeType(t);
	if (type === undefined) {
		return 'invalid';
	}
	if (b instanceof Uint8Array && b.length > 0 && b.length <= MAX_SHARED_MEDIA_BYTES) {
		return { t: type, b };
	}
	return b === undefined && typeof n === 'number' ? { t: type, n } : 'invalid';
};

const entrySize = (value: unknown): number => {
	const entry = readEntry(value);
	return typeof entry === 'object' ? (entry.b?.length ?? 0) : 0;
};

const isPlayable = (type: string): boolean => /^(?:audio|video)\//u.test(type);

/**
 * A URL of this page for `bytes`: a `blob:` URL for a raster picture, audio or video (types a
 * browser shows inertly when the URL is opened on its own), a `data:` URL for anything else,
 * whose origin is opaque.
 */
function localUrlFor(bytes: Uint8Array, type: string): string {
	const url = globalThis.URL as { createObjectURL?: (blob: Blob) => string } | undefined;
	if (
		typeof url?.createObjectURL === 'function' &&
		typeof globalThis.Blob === 'function' &&
		(mayUseOriginBlobUrl(type) || isPlayable(type))
	) {
		return url.createObjectURL(new Blob([bytes.slice()], { type }));
	}
	return `data:${type};base64,${encodeBase64(bytes)}`;
}

async function readLocalUrl(url: string): Promise<{ bytes: Uint8Array; type: string }> {
	const blob = await (await fetch(url)).blob();
	return {
		bytes: new Uint8Array(await blob.arrayBuffer()),
		type: normalizeMimeType(blob.type.split(';', 1)[0]) ?? 'application/octet-stream',
	};
}

const bridges = new WeakMap<YDocLike, CollaborationMedia>();

/** The media of `doc`'s room as this page sees them (one per document). */
export function collaborationMedia(doc: YDocLike): CollaborationMedia {
	let media = bridges.get(doc);
	if (!media) {
		media = createCollaborationMedia(doc);
		bridges.set(doc, media);
	}
	return media;
}

/** Set how `doc`'s room media are shared and shown in this viewer. */
export function configureCollaborationMedia(
	doc: YDocLike,
	options: CollaborationMediaOptions,
): void {
	collaborationMedia(doc).configure(options);
}

function createCollaborationMedia(doc: YDocLike): CollaborationMedia {
	const map = (): ObservableMap => doc.getMap(YDOC_MEDIA_KEY) as ObservableMap;
	const minted = new Map<string, { url: string; type: string }>();
	const mintedTypes = new Map<string, string>();
	const inflight = new Set<string>();
	// Media this page wrote into this room: read back as the URL it was written from.
	const written = new Map<string, string>();
	let options: CollaborationMediaOptions = {};

	const destroyable = doc as { on?: (event: string, handler: () => void) => void };
	destroyable.on?.('destroy', () => {
		for (const { url } of minted.values()) {
			forgetMediaUrl(url);
			if (isLocalOnlyUrl(url)) {
				globalThis.URL?.revokeObjectURL?.(url);
			}
		}
		minted.clear();
		mintedTypes.clear();
	});

	const put = (id: string, bytes: Uint8Array, type: string, origin: unknown): void => {
		doc.transact(() => {
			const entries = map();
			if (entries.get(id) !== undefined) {
				return;
			}
			const used = referencedMediaIds(doc);
			const unused: string[] = [];
			let total = 0;
			entries.forEach((value, key) => {
				if (used && !used.has(key)) {
					unused.push(key);
				} else {
					total += entrySize(value);
				}
			});
			for (const key of unused) {
				entries.delete(key);
			}
			entries.set(
				id,
				total + bytes.length > (options.maxBytes ?? MAX_SHARED_MEDIA_BYTES)
					? { t: type, n: bytes.length }
					: { t: type, b: bytes },
			);
		}, origin);
	};

	return {
		configure(next) {
			options = { ...options, ...next };
		},
		get allowExternalContent() {
			return options.allowExternalContent === true;
		},
		localMediaUrl(path) {
			const url = options.localMediaUrl?.(path);
			return isLocalOnlyUrl(url) ? url : undefined;
		},
		resolve(id) {
			const known = minted.get(id)?.url ?? written.get(id);
			if (known !== undefined) {
				return { url: known };
			}
			const entry = readEntry(map().get(id));
			if (entry === undefined) {
				return { missing: 'pending' };
			}
			if (entry === 'invalid') {
				return { missing: 'invalid' };
			}
			if (!entry.b) {
				return { missing: 'withheld' };
			}
			const url = localUrlFor(entry.b, entry.t);
			minted.set(id, { url, type: entry.t });
			mintedTypes.set(url, entry.t);
			rememberMediaUrl(url, id);
			return { url };
		},
		accepts(url, kind) {
			if (typeof url !== 'string') {
				return false;
			}
			const type = mintedTypes.get(url);
			if (type !== undefined) {
				return resourceKindAccepts(kind, type);
			}
			return isLocalOnlyUrl(url) && mediaIdOf(url) !== undefined;
		},
		ownsUrl(url) {
			return mediaIdOf(url) !== undefined;
		},
		async upload(ids, origin) {
			const pending = [...ids];
			for (const id of pending) {
				const url = mediaUrlOf(id);
				if (url !== undefined && !written.has(id)) {
					written.set(id, url);
				}
			}
			for (const id of pending) {
				const url = written.get(id);
				if (url === undefined || inflight.has(id) || map().get(id) !== undefined) {
					continue;
				}
				inflight.add(id);
				try {
					const { bytes, type } = await readLocalUrl(url);
					if (bytes.length > 0) {
						put(id, bytes, type, origin);
					}
				} catch {
					// The URL was revoked or cannot be read: there is nothing to share.
				} finally {
					inflight.delete(id);
				}
			}
		},
		observe(onChange) {
			const entries = map();
			const handler = (_event: unknown, transaction: YTransactionLike): void =>
				onChange(undefined, transaction);
			entries.observe?.(handler);
			return () => entries.unobserve?.(handler);
		},
	};
}
