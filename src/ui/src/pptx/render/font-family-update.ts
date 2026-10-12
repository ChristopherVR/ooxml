/**
 * What picking a font family in the ribbon or the inspector writes to a run.
 *
 * A run names up to four faces (`a:latin`, `a:ea`, `a:cs`, `a:sym`) and the
 * renderer paints each script in its own face (`text-script-fonts.ts`). The
 * font picker only ever set `fontFamily` (the `a:latin` face), so a run whose
 * `a:ea` named the same font as its `a:latin` kept painting Hangul, kana and
 * Han in the old face: on a Korean or Japanese slide choosing a font changed
 * nothing visible. A deck written by PowerPoint for a CJK locale, or by most
 * converters, gives every slot the same face, so this is the common case.
 *
 * The rule kept here is deliberately narrow: an east Asian or complex script
 * face follows the new family only when it was the same face as the old
 * `a:latin`, i.e. the deck treated the slots as one font. A run with a
 * distinct Asian face (`a:latin` Arial with `a:ea` Meiryo, or a theme whose
 * minor east Asian font differs from its minor Latin font) keeps it, since
 * the picker cannot tell whether the new font covers the script. PowerPoint
 * makes the same call: a Latin font picked from the font box leaves `a:ea`
 * alone.
 *
 * A concrete family also drops the theme token of each face it replaces. The
 * writer prefers the token (`+mn-lt`) over the concrete name, so leaving it in
 * place lost the choice on save.
 */

import type { TextStyle } from 'ooxml-core/pptx';

type ScriptFaceKey = 'eastAsiaFont' | 'complexScriptFont';
type ScriptTokenKey = 'eastAsiaFontThemeToken' | 'complexScriptFontThemeToken';

const SCRIPT_FACES: ReadonlyArray<{ face: ScriptFaceKey; token: ScriptTokenKey }> = [
	{ face: 'eastAsiaFont', token: 'eastAsiaFontThemeToken' },
	{ face: 'complexScriptFont', token: 'complexScriptFontThemeToken' },
];

/**
 * Expand a style update for `style`: when it sets a concrete `fontFamily`,
 * carry the change to the script faces that were the same face as the old
 * Latin one and drop the theme tokens of the faces it replaces. Any other
 * update comes back unchanged, and a script face the update names itself is
 * kept as given.
 */
export function fontFamilyStyleUpdates(
	style: TextStyle | undefined,
	updates: Partial<TextStyle>,
): Partial<TextStyle> {
	const family = updates.fontFamily;
	if (typeof family !== 'string' || family.length === 0) {
		return updates;
	}
	const next: Partial<TextStyle> = { ...updates };
	if (!('latinFontThemeToken' in updates) && style?.latinFontThemeToken !== undefined) {
		next.latinFontThemeToken = undefined;
	}
	const latinFace = style?.fontFamily;
	for (const { face, token } of SCRIPT_FACES) {
		if (face in updates || token in updates) {
			continue;
		}
		const current = style?.[face];
		if (current === undefined || current !== latinFace) {
			continue;
		}
		next[face] = family;
		if (style?.[token] !== undefined) {
			next[token] = undefined;
		}
	}
	return next;
}

/** Merge `updates` over `style`, applying {@link fontFamilyStyleUpdates}. */
export function mergeTextStyleUpdate(
	style: TextStyle | undefined,
	updates: Partial<TextStyle>,
): TextStyle {
	return { ...style, ...fontFamilyStyleUpdates(style, updates) };
}
