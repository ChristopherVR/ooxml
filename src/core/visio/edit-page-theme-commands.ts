import { fail } from './package-common';
import { isVisioBuiltInThemeId, type VisioBuiltInThemeId } from './theme-builtins';

/**
 * Design > Themes and Variants for one page. `theme` applies a built-in theme (writing its theme
 * part when the drawing lacks it) or `none` clears the page's theme selection; `variant` (0-3)
 * selects one of the theme's four colour and style variants. Without `theme`, only the variant of
 * the theme the page already uses changes.
 */
export interface VisioPageThemeEdit {
	type: 'set-page-theme';
	pageId: string;
	theme?: VisioBuiltInThemeId | 'none';
	variant?: number;
	/**
	 * Design > Variants > More > Colors: take only the colour scheme (ColorSchemeIndex, the
	 * variant selectors and the theme-colour caches) from a built-in theme, leaving the page's
	 * effect, connector and font scheme selections as they are. Excludes `theme`.
	 */
	colors?: VisioBuiltInThemeId;
	/** More > Fonts: only FontSchemeIndex, the theme fonts of themed text. */
	fonts?: VisioBuiltInThemeId;
	/** More > Effects: only EffectSchemeIndex, the theme effect styles of quick-styled shapes. */
	effects?: VisioBuiltInThemeId;
	/** More > Connectors: only ConnectorSchemeIndex, the theme line styles of connectors. */
	connectors?: VisioBuiltInThemeId;
}

/** The scheme parts that can be taken from a built-in theme on their own. */
export const PAGE_THEME_PARTS = ['colors', 'fonts', 'effects', 'connectors'] as const;

export function snapshotPageTheme(edit: VisioPageThemeEdit): VisioPageThemeEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	const result: VisioPageThemeEdit = { type: 'set-page-theme', pageId: edit.pageId };
	if (edit.theme !== undefined) {
		if (edit.theme !== 'none' && !isVisioBuiltInThemeId(edit.theme))
			fail('INVALID_EDIT', 'Unknown built-in theme.');
		result.theme = edit.theme;
	}
	const given = PAGE_THEME_PARTS.filter((part) => edit[part] !== undefined);
	if (given.length > 1 || (given.length && edit.theme !== undefined))
		fail('INVALID_EDIT', 'A theme or one scheme part can be applied at a time.');
	for (const part of given) {
		const value = edit[part];
		if (!isVisioBuiltInThemeId(value)) fail('INVALID_EDIT', 'Unknown built-in theme.');
		result[part] = value;
	}
	if (edit.variant !== undefined) {
		if (!Number.isInteger(edit.variant) || edit.variant < 0 || edit.variant > 3)
			fail('INVALID_EDIT', 'Theme variants are numbered 0 to 3.');
		if (edit.theme === 'none') fail('INVALID_EDIT', 'No Theme has no variants.');
		result.variant = edit.variant;
	}
	if (result.theme === undefined && !given.length && result.variant === undefined)
		fail('INVALID_EDIT', 'A page theme edit needs a theme, a scheme part or a variant.');
	return result;
}
