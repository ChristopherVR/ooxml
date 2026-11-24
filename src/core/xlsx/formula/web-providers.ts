// Opt-in fetch-backed default providers for `WEBSERVICE` and `IMAGE`.
// Because the calculation engine is synchronous, cache misses return `#BUSY!`, fetch in the
// background, store the settled result, and call `onUpdate` so the host can recalculate.
// Neither provider is enabled by default: a host must construct and pass it via
// `CalcEngineOptions.providers`. Disclosed in PRIVACY.md.
import type { CellError } from '../model';
import type { ImageProvider, ImageRequest, ImageSizing, WebserviceProvider } from './providers';
import { ERR, type Scalar } from './values';

const MAX_WEBSERVICE_CHARS = 32_767;
const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_MAX_IMAGE_BYTES = 5 * 1024 * 1024;

export type FetchLike = (input: string, init?: { signal?: AbortSignal }) => Promise<Response>;

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

export interface FetchImageProviderOptions extends FetchProviderBaseOptions {
	/** Maximum response byte length (default 5 MiB). */
	maxBytes?: number;
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
	const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
	let epoch = 0;

	return {
		resolve(request: ImageRequest): Scalar {
			const originErr = checkOrigin(request.source, options.allowedOrigins);
			if (originErr) return originErr;
			const cached = cache.get(request.source);
			if (cached !== undefined) {
				if ('error' in cached) return cached;
				return request.altText || cached.altText || request.source;
			}
			const fetcher =
				options.fetch ?? (typeof globalThis.fetch === 'function' ? globalThis.fetch : undefined);
			if (!fetcher) return ERR.CONNECT;
			if (!inflight.has(request.source)) {
				inflight.add(request.source);
				const myEpoch = epoch;
				void runWithTimeout(timeoutMs, (signal) =>
					fetcher(request.source, signal ? { signal } : undefined),
				)
					.then(async (res): Promise<LoadedCellImage | CellError> => {
						if (!res.ok) return ERR.CONNECT;
						const mimeType = (res.headers.get('content-type') ?? '')
							.split(';')[0]!
							.trim()
							.toLowerCase();
						if (!mimeType.startsWith('image/')) return ERR.VALUE;
						const buf = new Uint8Array(await res.arrayBuffer());
						if (buf.byteLength === 0 || buf.byteLength > maxBytes) return ERR.VALUE;
						return {
							url: request.source,
							mimeType,
							bytes: buf,
							altText: request.altText,
							sizing: request.sizing,
							...(request.height !== undefined ? { height: request.height } : {}),
							...(request.width !== undefined ? { width: request.width } : {}),
						};
					})
					.catch(() => ERR.CONNECT)
					.then((settled) => {
						inflight.delete(request.source);
						if (epoch !== myEpoch) return;
						cache.set(request.source, settled);
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
			cache.set(url, image);
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
