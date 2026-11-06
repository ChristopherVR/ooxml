/**
 * Fetches the actual font-file bytes for a Google Fonts catalogue webfont, so
 * `text-warp-outline-font-cache.ts` can parse a real outline for it.
 *
 * `google-webfonts.ts` already resolves which catalogue family to load and
 * injects a `<link rel=stylesheet>` for on-screen text, but a CSS stylesheet
 * carries no glyph geometry: only a `src: url(...)` pointing at the actual
 * `.woff2` file does. This module re-parses that same stylesheet text to find
 * the `@font-face` block matching the requested family/weight/style, then
 * fetches ITS `url()` for the binary.
 *
 * The `<link>`'s copy of the stylesheet cannot answer this `fetch`: Google
 * sends it with `Vary: Sec-Fetch-Dest, Sec-Fetch-Mode`, and a stylesheet
 * load (`style`, `no-cors`) and a script fetch (`empty`, `cors`) differ in
 * both. So each stylesheet and each font file is fetched once per fetch
 * implementation and the promise is shared (see {@link fetchOnce}): every
 * family and variant of one deck reads the same stylesheet, and concurrent
 * requests would otherwise all miss the HTTP cache together.
 *
 * Best-effort throughout: Google's CSS2 API can return several `@font-face`
 * blocks per family (subset by `unicode-range`, e.g. separate blocks for
 * `latin`, `latin-ext`, `cyrillic`, ...); this picks the block whose
 * `unicode-range` covers Basic Latin (`U+00..`, present in every WordArt
 * caption in practice) or, failing that, the first matching block. Any
 * failure (network, parse, no match) resolves to `undefined`; the caller
 * simply keeps the existing affine-transform fallback for that glyph.
 */
import { resolveFontItalic, resolveFontWeight, resolveGoogleFontsFace } from './font-weight-names';
import { findMetricCompatibleGoogleFontsFamily } from './google-webfonts-metric-clones';

interface ParsedFontFaceBlock {
	family: string;
	weight: number;
	italic: boolean;
	url: string;
	unicodeRange?: string;
}

function unquote(value: string): string {
	return value.trim().replace(/^['"]|['"]$/gu, '');
}

/** Parse every `@font-face { ... }` block out of a CSS2 stylesheet response. */
function parseFontFaceBlocks(css: string): ParsedFontFaceBlock[] {
	const blocks: ParsedFontFaceBlock[] = [];
	const blockRe = /@font-face\s*\{([^}]*)\}/gu;
	let match: RegExpExecArray | null;
	while ((match = blockRe.exec(css)) !== null) {
		const body = match[1];
		const familyMatch = /font-family:\s*([^;]+);/u.exec(body);
		const weightMatch = /font-weight:\s*([0-9]+)/u.exec(body);
		const styleMatch = /font-style:\s*(\w+)/u.exec(body);
		const urlMatch = /src:\s*url\(([^)]+)\)/u.exec(body);
		const rangeMatch = /unicode-range:\s*([^;]+);/u.exec(body);
		if (!familyMatch || !weightMatch || !urlMatch) {
			continue;
		}
		blocks.push({
			family: unquote(familyMatch[1]),
			weight: Number(weightMatch[1]),
			italic: styleMatch?.[1] === 'italic',
			url: unquote(urlMatch[1]),
			unicodeRange: rangeMatch?.[1]?.trim(),
		});
	}
	return blocks;
}

/** Basic Latin (U+0000-00FF) is the range covering ordinary caption text. */
function coversBasicLatin(unicodeRange: string | undefined): boolean {
	return !unicodeRange || /U\+00/iu.test(unicodeRange);
}

function pickBlock(
	blocks: readonly ParsedFontFaceBlock[],
	family: string,
	targetWeight: number,
	italic: boolean,
): ParsedFontFaceBlock | undefined {
	const byFamily = blocks.filter((b) => b.family.toLowerCase() === family.toLowerCase());
	const pool = byFamily.length > 0 ? byFamily : blocks;
	const byStyle = pool.filter((b) => b.weight === targetWeight && b.italic === italic);
	const candidates = byStyle.length > 0 ? byStyle : pool.filter((b) => b.italic === italic);
	if (candidates.length === 0) {
		return pool[0];
	}
	return candidates.find((b) => coversBasicLatin(b.unicodeRange)) ?? candidates[0];
}

/** `fetch`-shaped dependency, injected so this stays testable without a real network. */
export type FetchLike = (
	url: string,
) => Promise<{ ok: boolean; text: () => Promise<string>; arrayBuffer: () => Promise<ArrayBuffer> }>;

type CachedResponse = string | Uint8Array | undefined;

/** In-flight and settled fetches per fetch implementation, keyed by `kind:url`. */
const fetchCaches = new WeakMap<FetchLike, Map<string, Promise<CachedResponse>>>();

/**
 * Fetch `url` once per `fetchImpl` and share the promise. A failure is not
 * kept, so a later call can retry (after going offline, for example).
 */
function fetchOnce(fetchImpl: FetchLike, url: string, kind: 'text'): Promise<string | undefined>;
function fetchOnce(
	fetchImpl: FetchLike,
	url: string,
	kind: 'bytes',
): Promise<Uint8Array | undefined>;
function fetchOnce(
	fetchImpl: FetchLike,
	url: string,
	kind: 'text' | 'bytes',
): Promise<CachedResponse> {
	let cache = fetchCaches.get(fetchImpl);
	if (!cache) {
		cache = new Map();
		fetchCaches.set(fetchImpl, cache);
	}
	const key = `${kind}:${url}`;
	const cached = cache.get(key);
	if (cached) {
		return cached;
	}
	const pending = (async (): Promise<CachedResponse> => {
		try {
			const response = await fetchImpl(url);
			if (!response.ok) {
				return undefined;
			}
			return kind === 'text' ? await response.text() : new Uint8Array(await response.arrayBuffer());
		} catch {
			return undefined;
		}
	})();
	cache.set(key, pending);
	void pending.then((value) => {
		if (value === undefined) {
			cache.delete(key);
		}
	});
	return pending;
}

/**
 * Fetch the outline-ready bytes for `referencedFamily` (the family as
 * referenced in the deck, e.g. "Calibri") at `cssHref` (the same href
 * `resolveGoogleWebfontHref` built), resolving through the same
 * canonical/metric-clone lookup `matchGoogleWebfontFragments` used to decide
 * what the stylesheet actually serves.
 *
 * Returns `undefined` on any failure; never throws.
 */
export async function fetchGoogleWebfontOutlineBytes(
	cssHref: string,
	referencedFamily: string,
	bold: boolean,
	italic: boolean,
	fetchImpl: FetchLike,
): Promise<Uint8Array | undefined> {
	const canonical =
		resolveGoogleFontsFace(referencedFamily)?.family ??
		findMetricCompatibleGoogleFontsFamily(referencedFamily);
	if (!canonical) {
		return undefined;
	}
	try {
		const css = await fetchOnce(fetchImpl, cssHref, 'text');
		if (css === undefined) {
			return undefined;
		}
		// A weight variant ("DM Sans Light") is the base family's face at the
		// weight its name implies, exactly as the text itself renders.
		const weight = resolveFontWeight(referencedFamily, bold);
		const block = pickBlock(
			parseFontFaceBlocks(css),
			canonical,
			weight,
			resolveFontItalic(referencedFamily, italic),
		);
		if (!block) {
			return undefined;
		}
		return await fetchOnce(fetchImpl, block.url, 'bytes');
	} catch {
		return undefined;
	}
}
