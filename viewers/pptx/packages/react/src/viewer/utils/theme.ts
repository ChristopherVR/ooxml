/**
 * Thin re-export shim. The framework-agnostic PPTX document-theme font/colour
 * resolution helpers now live in `pptx-viewer-shared`.
 */
export {
	resolveThemeFont,
	tintColor,
	shadeColor,
	THEME_COLOR_TINT_ROWS,
	THEME_COLOR_LABELS,
	themeColorLabel,
	buildThemeColorGrid,
	themeColorSchemeToSwatches,
} from 'ooxml-ui/pptx';
export type { ThemeColorTintRow } from 'ooxml-ui/pptx';
