/** Compatibility entry: the implementation lives in ooxml-core/pptx/ui (`theme-color-swatches.ts`). */
export {
	buildThemeColorSwatchGrid,
	describeThemeColorSwatch,
	THEME_COLOR_SWATCH_COLUMN_LABEL_KEYS,
	THEME_COLOR_SWATCH_VARIANT_LABEL_KEYS,
	themeColorVariantOfRef,
	themeColorVariantToRef,
	themeColorVariantsForLuminance,
} from 'ooxml-core/pptx/ui';
export type {
	ThemeColorSwatch,
	ThemeColorSwatchColumn,
	ThemeColorVariant,
} from 'ooxml-core/pptx/ui';
