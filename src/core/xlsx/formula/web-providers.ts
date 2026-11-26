// Opt-in fetch-backed default providers for `WEBSERVICE` and `IMAGE`.
// Because the calculation engine is synchronous, cache misses return `#BUSY!`, fetch in the
// background, store the settled result, and call `onUpdate` so the host can recalculate.
// Neither provider is enabled by default: a host must construct and pass it via
// `CalcEngineOptions.providers`. Disclosed in PRIVACY.md.
import { RASTER_IMAGE_LIMITS } from '../../drawingml/raster-image';
import type { CellError } from '../model';
import { cellPictureInfo } from '../rich-values';
import type { ImageProvider, ImageRequest, ImageSizing, WebserviceProvider } from './providers';
import { ERR, type Scalar } from './values';

const MAX_WEBSERVICE_CHARS = 32_767;
const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_MAX_IMAGE_BYTES = 5 * 1024 * 1024;

export type FetchLike = (
	input: string,
	init?: Pick<RequestInit, 'signal' | 'redirect' | 'credentials' | 'referrerPolicy'>,
) => Promise<Response>;

export interface FetchProviderBaseOptions {
	/** Custom fetch implementation (defaults to `globalThis.fetch`). */
	fetch?: FetchLike;
	/** Optional list of allowed origins (e.g. `['https://example.com']`). Disallowed origins return `#BLOCKED!`. */
	allowedOrigins?: readonly string[];
	/** Request timeout in milliseconds (default 10,000). */
	timeoutMs?: number;
	/** Called when an in-flight request settles so the host can trigger recalculation. */
	onUpdate?: (url: string) => void;
}

export interface FetchWebserviceProviderOptions extends FetchProviderBaseOptions {
	/** Maximum UTF-16 response length before returning `#VALUE!` (default 32,767). */
	maxChars?: number;
}

export interface CachingWebserviceProvider extends WebserviceProvider {
	/** Pre-populates or overrides a cached URL response. */
	seed(url: string, value: string | CellError): void;
	/** Clears all cached entries and forgets in-flight tracking. */
	clear(): void;
	/** Number of currently in-flight fetches. */
	pending(): number;
}

export interface LoadedCellImage {
	readonly url: string;
	readonly mimeType: string;
	readonly bytes: Uint8Array;
	readonly altText: string;
	readonly sizing: ImageSizing;
	readonly height?: number;
	readonly width?: number;
}

export interface FetchImageProviderOptions extends Omit<FetchProviderBaseOptions, 'fetch'> {
	/** The host's HTTPS fetcher. Omitting it permits only explicitly seeded images. */
	fetch?: FetchLike;
	/** Maximum response byte length (default 5 MiB). */
	maxBytes?: number;
}

async function readImageBody(
	response: Response,
	limit: number,
	signal: AbortSignal | undefined,
): Promise<Uint8Array | undefined> {
	const reader = response.body?.getReader();
	if (!reader) return undefined;
	let bytes = new Uint8Array(0);
	let size = 0;
	const cancel = () => void reader.cancel().catch(() => undefined);
	signal?.addEventListener('abort', cancel, { once: true });
	try {
		signal?.throwIfAborted();
		if (Number(response.headers.get('content-length')) > limit) return undefined;
		for (let chunks = 0; chunks < 16_384; chunks++) {
			const { done, value } = await reader.read();
			signal?.throwIfAborted();
			if (done) return bytes.slice(0, size);
			const next = size + value.byteLength;
			if (next > limit) return undefined;
			if (next > bytes.byteLength) {
				const grown = new Uint8Array(Math.min(limit, Math.max(next, 4096, bytes.byteLength * 2)));
				grown.set(bytes);
				bytes = grown;
			}
			bytes.set(value, size);
			size = next;
		}
		return undefined;
	} finally {
		signal?.removeEventListener('abort', cancel);
		cancel();
		reader.releaseLock();
	}
}

export interface CachingImageProvider extends ImageProvider {
	/** Returns the loaded image bytes & metadata for a resolved URL, if available. */
	getImage(url: string): LoadedCellImage | undefined;
	/** Pre-populates or overrides a cached image entry. */
	seed(url: string, image: LoadedCellImage | CellError): void;
	/** Clears all cached entries and forgets in-flight tracking. */
	clear(): void;
	/** Number of currently in-flight fetches. */
	pending(): number;
}

function checkOrigin(url: string, allowedOrigins: readonly string[] | undefined): CellError | null {
	let parsed: URL;
	try {
		parsed = new URL(url);
	} catch {
		return ERR.VALUE;
	}
	if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return ERR.BLOCKED;
	if (allowedOrigins) {
		const allowed = new Set(allowedOrigins.map((o) => o.toLowerCase()));
		if (!allowed.has(parsed.origin.toLowerCase())) return ERR.BLOCKED;
	}
	return null;
}

async function runWithTimeout<T>(
	timeoutMs: number,
	fn: (signal: AbortSignal | undefined) => Promise<T>,
): Promise<T> {
	if (typeof AbortController === 'undefined' || !(timeoutMs > 0)) return fn(undefined);
	const controller = new AbortController();
	const timer = setTimeout(() => controller.abort(), timeoutMs);
	try {
		return await fn(controller.signal);
	} finally {
		clearTimeout(timer);
	}
}

/**
 * Creates an opt-in fetch-backed `WebserviceProvider` with an async cache.
 * Uncached URLs return `#BUSY!` while fetching and invoke `onUpdate(url)` when settled.
 */
export function createFetchWebserviceProvider(
	options: FetchWebserviceProviderOptions = {},
): CachingWebserviceProvider {
	const cache = new Map<string, string | CellError>();
	const inflight = new Set<string>();
	const maxChars = options.maxChars ?? MAX_WEBSERVICE_CHARS;
	const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
	let epoch = 0;

	return {
		fetch(url: string): string | CellError {
			const originErr = checkOrigin(url, options.allowedOrigins);
			if (originErr) return originErr;
			const cached = cache.get(url);
			if (cached !== undefined) return cached;
			const fetcher =
				options.fetch ?? (typeof globalThis.fetch === 'function' ? globalThis.fetch : undefined);
			if (!fetcher) return ERR.CONNECT;
			if (!inflight.has(url)) {
				inflight.add(url);
				const myEpoch = epoch;
				void runWithTimeout(timeoutMs, (signal) => fetcher(url, signal ? { signal } : undefined))
					.then(async (res) => {
						if (!res.ok) return ERR.CONNECT;
						const text = await res.text();
						return text.length > maxChars ? ERR.VALUE : text;
					})
					.catch(() => ERR.CONNECT)
					.then((settled) => {
						inflight.delete(url);
						if (epoch !== myEpoch) return;
						cache.set(url, settled);
						options.onUpdate?.(url);
					});
			}
			return ERR.BUSY;
		},
		seed(url, value) {
			cache.set(url, value);
		},
		clear() {
			epoch++;
			cache.clear();
			inflight.clear();
		},
		pending() {
			return inflight.size;
		},
	};
}

/**
 * Creates an opt-in fetch-backed `ImageProvider` with an async cache.
 * Uncached image URLs return `#BUSY!` while fetching and invoke `onUpdate(url)` when settled.
 */
export function createFetchImageProvider(
	options: FetchImageProviderOptions = {},
): CachingImageProvider {
	const cache = new Map<string, LoadedCellImage | CellError>();
	const inflight = new Set<string>();
	const maxBytes = options.maxBytes ?? DEFAULT_MAX_IMAGE_BYTES;
	if (
		!Number.isSafeInteger(maxBytes) ||
		maxBytes <= 0 ||
		maxBytes > RASTER_IMAGE_LIMITS.maxImageBytes
	)
		throw new RangeError('Invalid IMAGE byte limit.');
	const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
	let epoch = 0;
	let cachedBytes = 0;
	const store = (url: string, entry: LoadedCellImage | CellError): void => {
		const prior = cache.get(url);
		if (prior && !('error' in prior)) cachedBytes -= prior.bytes.byteLength;
		if ('error' in entry) cache.set(url, entry);
		else if (cachedBytes + entry.bytes.byteLength > 32 * 1024 * 1024) cache.set(url, ERR.BLOCKED);
		else {
			cachedBytes += entry.bytes.byteLength;
			cache.set(url, entry);
		}
	};

	return {
		resolve(request: ImageRequest): Scalar {
			const originErr = checkOrigin(request.source, options.allowedOrigins);
			if (originErr) return originErr;
			const url = new URL(request.source);
			if (url.protocol !== 'https:' || url.username || url.password) return ERR.BLOCKED;
			const cached = cache.get(request.source);
			if (cached !== undefined) {
				if ('error' in cached) return cached;
				return request.altText || cached.altText || request.source;
			}
			const fetcher = options.fetch;
			if (!fetcher) return ERR.BLOCKED;
			if (!inflight.has(request.source)) {
				if (cache.size + inflight.size >= 128 || inflight.size >= 8) return ERR.BLOCKED;
				inflight.add(request.source);
				const myEpoch = epoch;
				void runWithTimeout(timeoutMs, async (signal): Promise<LoadedCellImage | CellError> => {
					const res = await fetcher(request.source, {
						redirect: 'error',
						credentials: 'omit',
						referrerPolicy: 'no-referrer',
						...(signal ? { signal } : {}),
					});
					if (res.redirected) return ERR.BLOCKED;
					if (!res.ok) return ERR.CONNECT;
					const buf = await readImageBody(res, maxBytes, signal);
					if (!buf) return ERR.BLOCKED;
					const info = cellPictureInfo(buf);
					if (!info) return ERR.VALUE;
					return {
						url: request.source,
						mimeType: info.mimeType,
						bytes: buf,
						altText: request.altText,
						sizing: request.sizing,
						...(request.height !== undefined ? { height: request.height } : {}),
						...(request.width !== undefined ? { width: request.width } : {}),
					};
				})
					.catch(() => ERR.CONNECT)
					.then((settled) => {
						if (epoch !== myEpoch) return;
						inflight.delete(request.source);
						store(request.source, settled);
						options.onUpdate?.(request.source);
					});
			}
			return ERR.BUSY;
		},
		getImage(url) {
			const entry = cache.get(url);
			return entry && !('error' in entry) ? entry : undefined;
		},
		seed(url, image) {
			if (!cache.has(url) && cache.size + inflight.size >= 128) return;
			if ('error' in image) store(url, image);
			else {
				const bytes = image.bytes.byteLength <= maxBytes ? image.bytes.slice() : undefined;
				const info = bytes && cellPictureInfo(bytes);
				store(url, bytes && info ? { ...image, bytes, mimeType: info.mimeType } : ERR.VALUE);
			}
		},
		clear() {
			epoch++;
			cache.clear();
			inflight.clear();
			cachedBytes = 0;
		},
		pending() {
			return inflight.size;
		},
	};
}
