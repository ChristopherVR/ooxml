// Compatibility exports: document operations live in ooxml-core.
export {
	resolveThemeFont,
	tintColor,
	shadeColor,
	THEME_COLOR_TINT_ROWS,
	THEME_COLOR_LABELS,
	buildThemeColorGrid,
	themeColorLabel,
	themeColorSchemeToSwatches,
} from 'ooxml-core/pptx/editor/render/text-theme';
export type { ThemeColorTintRow } from 'ooxml-core/pptx/editor/render/text-theme';
