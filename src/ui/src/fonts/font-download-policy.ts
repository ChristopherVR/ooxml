/**
 * Who decides whether missing fonts are downloaded: the host's `googleFonts` value when it set
 * one, otherwise the reader's Trust Center choice ("Download fonts that are not installed or
 * embedded"). Shared by Word, Excel and Visio, whose Trust Center toggle is off until the reader
 * turns it on (they made no network request before it existed); PowerPoint keeps its own policy
 * (`pptx/render/options/font-download-policy.ts`), whose toggle starts on.
 *
 * @module fonts/font-download-policy
 */
import {
	GOOGLE_FONTS_CSS2_BASE,
	resolveGoogleFontsCss2Base,
	type GoogleFontsOption,
} from 'ooxml-core/text';

/** The Trust Center toggle's default for Word, Excel and Visio. */
export const FONT_DOWNLOADS_DEFAULT = false;

/**
 * The option the product's webfont loader uses: the host's value when set, else Google Fonts when
 * the reader allowed downloads and `false` when not.
 */
export function effectiveFontDownloads(
	host: GoogleFontsOption | undefined,
	readerAllows: boolean,
): GoogleFontsOption {
	if (host !== undefined) return host;
	return readerAllows ? { baseUrl: GOOGLE_FONTS_CSS2_BASE } : false;
}

/**
 * How the host's value locks the Trust Center toggle: `undefined` when the host left it to the
 * reader, otherwise the value the toggle shows read-only (on for a usable server, off for `false`
 * or an invalid `baseUrl`, which fails closed).
 */
export function fontDownloadLock(host: GoogleFontsOption | undefined): boolean | undefined {
	return host === undefined ? undefined : resolveGoogleFontsCss2Base(host, null, 'ooxml') !== null;
}
