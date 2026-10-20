/**
 * embedded-font-bytes.ts: validate an embedded font's URL and recover its
 * clear-text bytes. Split out of `embedded-fonts.ts` so the trust check in
 * `embedded-font-trust.ts` can read font files without a circular import.
 *
 * @module render/embedded-font-bytes
 */
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';
import { deobfuscateFont } from 'ooxml-core/pptx';

/** Whitelist of `data:`/`blob:` URLs safe to interpolate into `url("…")`. */
const FONT_DATA_URL_PATTERN =
	/^data:font\/[a-z0-9+.-]+(?:;charset=[a-z0-9-]+)?;base64,[a-z0-9+/=]+$/iu;

/** Decode a validated `data:font/…;base64,…` URL's payload to raw bytes. */
function decodeFontDataUrl(dataUrl: string): Uint8Array | undefined {
	const match = FONT_DATA_URL_PATTERN.exec(dataUrl);
	if (!match) {
		return undefined;
	}
	const commaIndex = dataUrl.indexOf(',');
	if (commaIndex < 0) {
		return undefined;
	}
	try {
		const binary = atob(dataUrl.slice(commaIndex + 1));
		const bytes = new Uint8Array(binary.length);
		for (let i = 0; i < binary.length; i++) {
			bytes[i] = binary.charCodeAt(i);
		}
		return bytes;
	} catch {
		return undefined;
	}
}

/**
 * True when `url` is safe to interpolate into `src: url("…")`.
 *
 * Accepts `blob:` URLs (minted by the binding from de-obfuscated bytes) and
 * base64 `data:font/…` URLs produced by the core loader; everything else is
 * rejected to prevent CSS-injection / exfiltration via a hostile PPTX.
 */
export function isInjectableUrl(url: string): boolean {
	if (typeof url !== 'string' || url.length === 0) {
		return false;
	}
	if (url.startsWith('blob:')) {
		return true;
	}
	return FONT_DATA_URL_PATTERN.test(url);
}

/**
 * Resolve an embedded-font entry to its raw (clear-text) font bytes,
 * independent of how the caller will use them (CSS `@font-face` injection via
 * {@link resolveFontVariant}, or glyph-outline parsing via
 * `text-warp-outline-font-cache.ts`). Tries, in order: decoding the core
 * loader's own `dataUrl` (the common case), the preserved `rawFontData`, and
 * de-obfuscating `originalPartBytes` with `fontGuid`. Returns `undefined` when
 * none of these yield usable bytes.
 */
export function resolveEmbeddedFontClearBytes(font: PptxEmbeddedFont): Uint8Array | undefined {
	if (isInjectableUrl(font.dataUrl) && !font.dataUrl.startsWith('blob:')) {
		const decoded = decodeFontDataUrl(font.dataUrl);
		if (decoded && decoded.length >= 4) {
			return decoded;
		}
	}
	if (font.rawFontData && font.rawFontData.length > 0) {
		return font.rawFontData;
	}
	if (font.originalPartBytes && font.originalPartBytes.length > 0 && font.fontGuid) {
		return deobfuscateFont(font.originalPartBytes, font.fontGuid);
	}
	return undefined;
}
