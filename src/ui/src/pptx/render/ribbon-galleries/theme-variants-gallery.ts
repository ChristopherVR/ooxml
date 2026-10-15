/**
 * Design > Variants: four colour variants of the current theme, plus the Effects and Background
 * Styles entries.
 *
 * A variant is the theme's own colour scheme read four ways: as it is, with light and dark
 * inverted (a dark background), and with the six accents shifted by two and by four places. They
 * are computed from the scheme, so they follow any theme, and a pick writes a colour scheme only
 * (fonts and the theme name stay), through the same path the Colors gallery uses.
 *
 * Theme effects and background styles have no editing engine behind them yet, so those two
 * entries are present, disabled, and say why in their tooltip.
 *
 * @module render/ribbon-galleries/theme-variants-gallery
 */
import type { PptxThemeColorScheme } from 'ooxml-core/pptx';

import { GALLERY_THEME_PRESETS } from '../../theme/theme-gallery-presets';
import type { RibbonGalleryModule } from './gallery-module';
import { svgTile } from './gallery-preview-svg';
import { colorSchemesEqual, galleryColorScheme } from './gallery-theme';
import type { RibbonGalleryContext, RibbonGalleryId } from './gallery-types';
import { themeTileSvg } from './themes-gallery';

const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;
const TILE = { width: 52, height: 30 };

export interface ThemeVariant {
	id: 'original' | 'inverted' | 'shiftTwo' | 'shiftFour';
	labelKey: string;
	label: string;
	scheme: PptxThemeColorScheme;
}

function shiftAccents(scheme: PptxThemeColorScheme, by: number): PptxThemeColorScheme {
	const next = { ...scheme };
	ACCENTS.forEach((key, index) => {
		next[key] = scheme[ACCENTS[(index + by) % ACCENTS.length]!];
	});
	return next;
}

/** The four variants of `base`, in gallery order. */
export function themeVariants(base: PptxThemeColorScheme): ThemeVariant[] {
	return [
		{
			id: 'original',
			labelKey: 'pptx.gallery.themeVariants.original',
			label: 'Variant 1: Original colors',
			scheme: { ...base },
		},
		{
			id: 'inverted',
			labelKey: 'pptx.gallery.themeVariants.inverted',
			label: 'Variant 2: Light and dark inverted',
			scheme: { ...base, dk1: base.lt1, lt1: base.dk1, dk2: base.lt2, lt2: base.dk2 },
		},
		{
			id: 'shiftTwo',
			labelKey: 'pptx.gallery.themeVariants.shiftTwo',
			label: 'Variant 3: Accent colors shifted by two',
			scheme: shiftAccents(base, 2),
		},
		{
			id: 'shiftFour',
			labelKey: 'pptx.gallery.themeVariants.shiftFour',
			label: 'Variant 4: Accent colors shifted by four',
			scheme: shiftAccents(base, 4),
		},
	];
}

/**
 * The scheme the variants derive from. When the deck's colours are one of a built-in theme's four
 * variants, that theme's own scheme is the base, so the four tiles stay put after a pick (whatever
 * the theme is called); any other colours are their own base.
 */
export function variantBase(ctx: RibbonGalleryContext): PptxThemeColorScheme {
	const current = galleryColorScheme(ctx);
	const preset = GALLERY_THEME_PRESETS.find((candidate) =>
		themeVariants(candidate.colorScheme).some((variant) =>
			colorSchemesEqual(variant.scheme, current),
		),
	);
	return preset?.colorScheme ?? current;
}

export const THEME_VARIANTS_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const current = galleryColorScheme(ctx);
		const heading = ctx.theme?.fontScheme?.majorFont?.latin;
		return {
			id: 'themeVariants',
			labelKey: 'pptx.gallery.themeVariants.title',
			label: 'Variants',
			disabled: false,
			inlineCount: 4,
			sections: [
				{
					id: 'variants',
					columns: 4,
					tileWidth: TILE.width,
					tileHeight: TILE.height,
					items: themeVariants(variantBase(ctx)).map((variant) => ({
						id: variant.id,
						labelKey: variant.labelKey,
						label: variant.label,
						previewSvg: themeTileSvg(variant.scheme, heading, TILE),
						applied: colorSchemesEqual(variant.scheme, current),
					})),
				},
			],
		};
	},
	apply(itemId, ctx) {
		const variant = themeVariants(variantBase(ctx)).find((candidate) => candidate.id === itemId);
		return variant
			? {
					kind: 'themeColorScheme',
					colorScheme: variant.scheme,
					name: ctx.theme?.name ?? 'Custom',
				}
			: null;
	},
};

/** A dropdown that is listed but has no engine behind it: disabled, with its reason. */
function unavailableModule(
	id: RibbonGalleryId,
	labelKey: string,
	label: string,
	iconPath: string,
	hint: { key: string; text: string },
): RibbonGalleryModule {
	return {
		build: () => ({
			id,
			labelKey,
			label,
			disabled: true,
			command: { iconPath, hintKey: hint.key, hint: hint.text },
			sections: [
				{
					id: 'unavailable',
					columns: 1,
					tileWidth: 0,
					tileHeight: 0,
					items: [
						{ id: 'run', labelKey, label, previewSvg: svgTile(0, 0, '', ''), applied: false },
					],
				},
			],
		}),
		apply: () => null,
	};
}

export const THEME_EFFECTS_GALLERY = unavailableModule(
	'themeEffects',
	'pptx.gallery.themeEffects.title',
	'Effects',
	'M10 2l2.4 5 5.4.6-4 3.7 1.1 5.4L10 14l-4.9 2.7 1.1-5.4-4-3.7 5.4-.6z',
	{
		key: 'pptx.gallery.themeEffects.unavailable',
		text: 'Theme effects cannot be changed in this viewer yet.',
	},
);

export const THEME_BACKGROUNDS_GALLERY = unavailableModule(
	'themeBackgrounds',
	'pptx.gallery.themeBackgrounds.title',
	'Background Styles',
	'M3 4h14v12H3zM3 13l4-4 3 3 3-3 4 4',
	{
		key: 'pptx.gallery.themeBackgrounds.unavailable',
		text: 'Background styles are not available in this viewer yet. Use Format Background to set the slide background.',
	},
);
