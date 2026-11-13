/**
 * Font substitution for PowerPoint: the shared Office chain (`ooxml-core/text`), with the
 * measured look-alike of a font Google Fonts does not serve (Aptos: Gantari) put straight after
 * the deck font.
 *
 * Only PowerPoint does this, because only PowerPoint corrects the stand-in's widths itself:
 * `text-design-tracking.ts` (`ooxml-ui/pptx`) spreads the deck font's measured advances over
 * whatever face the browser paints, so the stand-in can be chosen by its shapes rather than by
 * its metrics (issue #249). Word, Excel and Visio keep the metric clone first, since the browser's
 * own measurement decides their line breaks. The look-alike is only painted once the webfont
 * fallback has loaded it (or it is installed); otherwise the chain falls through to Calibri as
 * before.
 */
import { stripFontStyleSuffix } from '../../../text/font-style-name';
import {
	buildFontFamilyString,
	getSubstituteFonts as getSharedSubstituteFonts,
} from '../../../text/font-substitution';
import { findLookAlikeGoogleFontsFamily } from '../../../text/webfonts/look-alikes';

export * from '../../../text/font-substitution';

/** The fallback fonts for `fontName` (without it), as PowerPoint orders them. */
export function getSubstituteFonts(fontName: string, panose?: number[]): readonly string[] {
	const fallbacks = getSharedSubstituteFonts(fontName, panose);
	const trimmed = fontName.trim();
	const base = stripFontStyleSuffix(trimmed);
	const lookAlike =
		findLookAlikeGoogleFontsFamily(trimmed) ?? (base ? findLookAlikeGoogleFontsFamily(base) : null);
	if (!lookAlike || fallbacks.includes(lookAlike)) {
		return fallbacks;
	}
	// A full name ("Aptos SemiBold") lists its base family first; the look-alike follows it.
	const at = base && fallbacks[0] === base ? 1 : 0;
	return [...fallbacks.slice(0, at), lookAlike, ...fallbacks.slice(at)];
}

/** The CSS `font-family` value for `fontName`, as PowerPoint orders the fallbacks. */
export function getSubstituteFontFamily(fontName: string, panose?: number[]): string {
	const trimmed = fontName.trim();
	if (!trimmed) {
		return 'sans-serif';
	}
	return buildFontFamilyString(trimmed, getSubstituteFonts(trimmed, panose));
}
