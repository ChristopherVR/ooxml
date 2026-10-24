/**
 * Design > Variants > Effects: the built-in theme effect schemes.
 *
 * An effect scheme is the theme's `a:fmtScheme`: the three fill styles, three outline styles,
 * three effect styles and three background fill styles every `p:style` reference
 * (`fillRef`/`lnRef`/`effectRef` idx 1 to 3, `bgRef` idx 1001 to 1003) points into. Applying one
 * replaces that element in the theme part, so every shape that references the theme redraws.
 *
 * The fragments are the ones PowerPoint writes into a deck when the matching `.eftx` is loaded
 * (see `theme-effect-schemes-data.ts`), so a saved deck matches what PowerPoint produces.
 *
 * @module utils/theme-effect-schemes
 */
import type { PptxThemeFormatScheme } from '../types';
import { THEME_EFFECT_SCHEME_SOURCES } from './theme-effect-schemes-data';

/** One built-in effect scheme. */
export interface PptxThemeEffectScheme {
	/** Stable id (`office`, `office2007`, `subtleSolids`, ...). */
	id: string;
	/** English name; also the `a:fmtScheme/@name` PowerPoint writes. */
	name: string;
	/** The `a:fmtScheme` element, with the `a` prefix undeclared (see {@link wrapFormatSchemeXml}). */
	fmtSchemeXml: string;
}

/** The built-in effect schemes, in the order of PowerPoint's Effects gallery. */
export const THEME_EFFECT_SCHEMES: readonly PptxThemeEffectScheme[] =
	THEME_EFFECT_SCHEME_SOURCES.map(([id, name, fmtSchemeXml]) => ({ id, name, fmtSchemeXml }));

/** The scheme with this id, or `undefined`. */
export function findThemeEffectScheme(id: string): PptxThemeEffectScheme | undefined {
	return THEME_EFFECT_SCHEMES.find((scheme) => scheme.id === id);
}

/** The scheme whose name is `name` (a deck's `a:fmtScheme/@name`), or `undefined`. */
export function findThemeEffectSchemeByName(
	name: string | undefined,
): PptxThemeEffectScheme | undefined {
	if (!name) {
		return undefined;
	}
	const wanted = name.trim().toLowerCase();
	return THEME_EFFECT_SCHEMES.find((scheme) => scheme.name.toLowerCase() === wanted);
}

/**
 * A format scheme's identity: its name and the raw XML of each style, attributes in name order and
 * child elements in document order (colour transforms apply in order). The colour a parsed `phClr`
 * style resolved to depends on the theme colours when it was read, so it is left out: two reads of
 * one `a:fmtScheme` have the same key, and so does a plain or reactive copy.
 */
export function formatSchemeKey(scheme: PptxThemeFormatScheme): string {
	const raw = (styles: readonly { rawNode?: unknown }[] | undefined) =>
		(styles ?? []).map((style) => style.rawNode ?? null);
	const lists = [scheme.fillStyles, scheme.lineStyles, scheme.effectStyles];
	return JSON.stringify(
		[scheme.name ?? null, ...lists.map(raw), raw(scheme.backgroundFillStyles)],
		(_key, value: unknown) => {
			if (!value || typeof value !== 'object' || Array.isArray(value)) {
				return value;
			}
			const entries = Object.entries(value);
			const isAttribute = ([key]: [string, unknown]) => key.startsWith('@_');
			const attributes = entries.filter(isAttribute).sort(([x], [y]) => (x < y ? -1 : 1));
			return Object.fromEntries([...attributes, ...entries.filter((entry) => !isAttribute(entry))]);
		},
	);
}

/** `fmtSchemeXml` with the DrawingML namespace declared, ready for an XML parser. */
export function wrapFormatSchemeXml(fmtSchemeXml: string): string {
	return fmtSchemeXml.replace(
		/^<a:fmtScheme/u,
		'<a:fmtScheme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	);
}
