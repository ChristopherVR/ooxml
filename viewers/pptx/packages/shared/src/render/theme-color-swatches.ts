/** Compatibility entry: the implementation lives in ooxml-core/pptx/ui (`theme-color-swatches.ts`). */
export {
	buildThemeColorSwatchGrid,
	describeThemeColorSwatch,
	themeColorVariantOfRef,
	themeColorVariantToRef,
	themeColorVariantsForLuminance,
} from 'pptx-viewer-core/ui';
export type {
	ThemeColorSwatch,
	ThemeColorSwatchColumn,
	ThemeColorVariant,
} from 'pptx-viewer-core/ui';
