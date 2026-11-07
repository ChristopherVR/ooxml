/**
 * Who decides whether the viewer downloads webfonts (the Google Fonts
 * fallback for families that are neither installed nor embedded): the host's
 * `googleFonts` prop / input / option, or the user's File > Options > Trust
 * Center setting `trust.allowFontDownloads`.
 *
 * - host unset: the Trust Center option decides (on requests Google Fonts,
 *   off requests nothing);
 * - host `false`: off, and the Trust Center toggle is locked off;
 * - host `{ baseUrl }` with a valid URL: on with that mirror, locked on;
 * - host `{ baseUrl }` with an invalid URL: off (it fails closed, as
 *   {@link resolveGoogleFontsCss2Base} does), locked off.
 *
 * Every binding feeds {@link resolveEffectiveGoogleFonts} (or the base URL
 * from {@link resolveFontDownloadBase}) to its webfont paths and folds
 * {@link lockFontDownloadSetting} into the options locks; neither rule is
 * re-derived per binding.
 *
 * @module render/options/font-download-policy
 */
import type { OptionsSettingId, OptionsSettingValues } from '../customization/customization-types';
import { resolveGoogleFontsCss2Base } from '../google-webfonts';
import type { GoogleFontsOption } from '../google-webfonts';
import type { ViewerOptions } from './viewer-options';

/** The `<group>.<key>` id of the Trust Center setting the host `googleFonts` value locks. */
export const FONT_DOWNLOAD_SETTING_ID: OptionsSettingId = 'trust.allowFontDownloads';

/**
 * The `googleFonts` policy a binding applies: the host value when set,
 * otherwise `false` when the user turned font downloads off in the Trust
 * Center, and `undefined` (Google Fonts) when it is on. Pass the result
 * wherever the host value went before; `resolveGoogleFontsCss2Base` of it is
 * {@link resolveFontDownloadBase}.
 */
export function resolveEffectiveGoogleFonts(
	host: GoogleFontsOption | undefined,
	options: Pick<ViewerOptions, 'trust'>,
): GoogleFontsOption | undefined {
	if (host !== undefined) {
		return host;
	}
	return options.trust.allowFontDownloads ? undefined : false;
}

/**
 * The CSS2 base URL the webfont fallback requests, or `null` when font
 * downloads are off (by the host, by an invalid mirror URL, or by the user).
 */
export function resolveFontDownloadBase(
	host: GoogleFontsOption | undefined,
	options: Pick<ViewerOptions, 'trust'>,
): string | null {
	return resolveGoogleFontsCss2Base(resolveEffectiveGoogleFonts(host, options));
}

/**
 * The value the host `googleFonts` setting forces on the Trust Center
 * toggle: `undefined` when unset (the user decides), otherwise whether the
 * host value lets fonts download.
 */
export function hostFontDownloadLock(host: GoogleFontsOption | undefined): boolean | undefined {
	return host === undefined ? undefined : resolveGoogleFontsCss2Base(host) !== null;
}

/**
 * The options locks with the host `googleFonts` value folded in. Returns
 * `locked` itself when the host left it unset, and a new record otherwise;
 * the host value wins over a `customization.options.locked` entry for the
 * same id.
 */
export function lockFontDownloadSetting(
	locked: OptionsSettingValues,
	host: GoogleFontsOption | undefined,
): OptionsSettingValues {
	const value = hostFontDownloadLock(host);
	if (value === undefined) {
		return locked;
	}
	return { ...locked, [FONT_DOWNLOAD_SETTING_ID]: value };
}
