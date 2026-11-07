/**
 * The Google Fonts CSS2 request every product builds for the fonts a document uses but the
 * reader has neither installed nor embedded: the host policy (`GoogleFontsOption`), the base URL
 * check, the per-family query fragment and the stylesheet href. Pure and DOM-free; the browser
 * side (installed check, `<link>` injection, waiting for faces) is `ooxml-ui`'s `fonts` folder.
 *
 * Which families are served is answered by the bundled catalogue (`google-fonts-catalogue.ts`),
 * never by probing the API, so the only family names that reach the server are public catalogue
 * names a document already uses. A family the catalogue does not serve under its own name
 * ("Calibri") is requested as its verified metric-compatible clone (`metric-clones.ts`), which
 * the product's CSS font stack already lists after the document font.
 *
 * Moved from PowerPoint's `src/ui/src/pptx/render/google-webfonts.ts` so Word, Excel and Visio
 * share it.
 *
 * @module text/webfonts/google-fonts-request
 */
import { resolveGoogleFontsFace } from './font-weight-names';
import { findMetricCompatibleGoogleFontsFamily } from './metric-clones';

/** Base URL of the Google Fonts CSS2 API. */
export const GOOGLE_FONTS_CSS2_BASE = 'https://fonts.googleapis.com/css2';

/**
 * Host policy for the webfont fallback, the `googleFonts` option of every product.
 *
 * - `false`: off. No stylesheet is injected, so no request reaches `fonts.googleapis.com` or
 *   `fonts.gstatic.com`; text uses installed, embedded or fallback fonts.
 * - `{ baseUrl }`: every CSS2 request goes to `baseUrl` instead, with the same query format
 *   (`?family=...&display=swap`).
 * - unset: each product's documented default (see {@link resolveGoogleFontsCss2Base}).
 */
export type GoogleFontsOption = false | { readonly baseUrl: string };

/**
 * Validate a self-hosted CSS2 mirror base URL: an absolute `http:` or `https:` URL without
 * credentials, query or fragment (the request appends its own query). The normalised URL, or
 * `null` when it is not acceptable.
 */
export function normalizeGoogleFontsBaseUrl(baseUrl: unknown): string | null {
	if (typeof baseUrl !== 'string' || baseUrl.trim() === '') return null;
	let url: URL;
	try {
		url = new URL(baseUrl.trim());
	} catch {
		return null;
	}
	if (url.protocol !== 'http:' && url.protocol !== 'https:') return null;
	if (url.username !== '' || url.password !== '' || url.search !== '' || url.hash !== '') {
		return null;
	}
	return url.href;
}

const warnedInvalidBaseUrls = new Set<string>();

/**
 * The CSS2 base URL a `googleFonts` value requests, or `null` when the fallback is off. An
 * invalid `baseUrl` fails closed: it warns (once per distinct value) and turns the fallback off
 * rather than silently falling back to Google. `whenUnset` is the product's default for an unset
 * value: PowerPoint's is Google (`GOOGLE_FONTS_CSS2_BASE`), Word, Excel and Visio's is off.
 */
export function resolveGoogleFontsCss2Base(
	option: GoogleFontsOption | undefined,
	whenUnset: string | null = GOOGLE_FONTS_CSS2_BASE,
	label = 'pptx-viewer',
): string | null {
	if (option === undefined) return whenUnset;
	if (option === false) return null;
	const raw = typeof option === 'object' && option !== null ? option.baseUrl : undefined;
	const normalized = normalizeGoogleFontsBaseUrl(raw);
	if (normalized === null && !warnedInvalidBaseUrls.has(String(raw ?? ''))) {
		warnedInvalidBaseUrls.add(String(raw ?? ''));
		console.warn(
			`[${label}] googleFonts.baseUrl must be an absolute http(s) URL without credentials, query or fragment; the Google Fonts fallback is disabled.`,
		);
	}
	return normalized;
}

/**
 * Axis spec requested for every catalogue family. The API serves only the weights and styles a
 * family has, so this one fragment is safe for every family and yields real bold and italic.
 */
export const GOOGLE_WEBFONT_AXIS = 'ital,wght@0,400;0,700;1,400;1,700';

/** `display=swap` so text paints with the fallback and swaps when ready. */
const DISPLAY_PARAM = 'display=swap';

/**
 * The referenced families to look up: everything not satisfied by an embedded font and (when
 * the check is supplied) not installed, so an installed family is used as-is and its name never
 * reaches the API.
 */
export function selectGoogleWebfontFamilies(
	referenced: Iterable<string>,
	embedded?: Iterable<string>,
	isLocallyInstalled?: (family: string) => boolean,
): string[] {
	const embeddedSet = new Set(embedded ?? []);
	const selected: string[] = [];
	for (const family of referenced) {
		if (embeddedSet.has(family) || isLocallyInstalled?.(family)) continue;
		selected.push(family);
	}
	return selected;
}

/**
 * The CSS2 query fragment for one family. Word separators stay literal spaces: the fragment goes
 * through `encodeURIComponent`, and an encoded `+` reaches Google as a literal plus it rejects.
 */
export function buildGoogleFontsFragment(family: string, weights?: Iterable<number>): string {
	const sorted = [...new Set([400, 700, ...(weights ?? [])])].sort((a, b) => a - b);
	if (sorted.length === 2) return `${family}:${GOOGLE_WEBFONT_AXIS}`;
	// The API wants the tuples sorted: every upright weight, then every italic.
	const tuples = [0, 1].flatMap((ital) => sorted.map((weight) => `${ital},${weight}`));
	return `${family}:ital,wght@${tuples.join(';')}`;
}

/** The CSS2 stylesheet href for the fragments, or `null` when there is nothing to load. */
export function buildGoogleFontsHref(
	fragments: readonly string[],
	baseUrl: string = GOOGLE_FONTS_CSS2_BASE,
): string | null {
	if (fragments.length === 0) return null;
	const query = fragments.map((fragment) => `family=${encodeURIComponent(fragment)}`).join('&');
	return `${baseUrl}?${query}&${DISPLAY_PARAM}`;
}

/** The family the stylesheet serves for `family` (itself or its clone), or `null`. */
export function servedGoogleFontsFamily(family: string): string | null {
	return resolveGoogleFontsFace(family)?.family ?? findMetricCompatibleGoogleFontsFamily(family);
}

/**
 * The query fragments for the candidate families the catalogue knows, under their canonical
 * spelling or as their metric-compatible clone; families resolving to one served family give one
 * fragment. A weight variant named as its own family ("DM Sans Light") requests its base family
 * with that weight on the axis. Unmatched families are dropped without any request. Sorted, so
 * the href depends only on which families a document uses. `onResolved` hears every candidate
 * that matched.
 */
export function matchGoogleWebfontFragments(
	families: readonly string[],
	onResolved?: (family: string) => void,
): string[] {
	const weightsByFamily = new Map<string, Set<number>>();
	for (const family of families) {
		const face = resolveGoogleFontsFace(family);
		const canonical = face?.family ?? findMetricCompatibleGoogleFontsFamily(family);
		if (canonical === null) continue;
		onResolved?.(family);
		let weights = weightsByFamily.get(canonical);
		if (!weights) {
			weights = new Set();
			weightsByFamily.set(canonical, weights);
		}
		if (face?.weight !== undefined) {
			// The regular face, and the face bold runs in this family ask for.
			weights.add(face.weight);
			weights.add(Math.max(700, face.weight));
		}
	}
	return [...weightsByFamily]
		.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
		.map(([family, weights]) => buildGoogleFontsFragment(family, weights));
}
