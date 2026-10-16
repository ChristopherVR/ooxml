/**
 * Design > Variants: the theme's variants.
 *
 * When the package carries real variants (`themeVariants/` parts, read at load into
 * `PptxData.themeVariants`) those are offered: a pick applies the variant's colours and fonts, its
 * effect scheme and its background style. Otherwise four colour variants are computed from the
 * theme's own colour scheme, and the gallery says so: the scheme as it is, with light and dark
 * inverted (a dark background), and with the six accents shifted by two and by four places. A
 * computed pick writes a colour scheme only (fonts and the theme name stay), through the same
 * path the Colors gallery uses.
 *
 * @module render/ribbon-galleries/theme-variants-gallery
 */
import type { PptxThemeColorScheme, PptxThemeVariant } from 'ooxml-core/pptx';
import { backgroundStylePreset } from 'ooxml-core/pptx';

import { GALLERY_THEME_PRESETS } from '../../theme/theme-gallery-presets';
import type { RibbonGalleryModule } from './gallery-module';
import { colorSchemesEqual, galleryColorScheme } from './gallery-theme';
import type { RibbonGalleryContext, RibbonGalleryItem } from './gallery-types';
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

/** The four computed variants of `base`, in gallery order. */
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
 * The scheme the computed variants derive from. When the deck's colours are one of a built-in
 * theme's four variants, that theme's own scheme is the base, so the four tiles stay put after a
 * pick (whatever the theme is called); any other colours are their own base.
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

/** The real variants the package carries, or none. */
function realVariants(ctx: RibbonGalleryContext): readonly PptxThemeVariant[] {
	return ctx.themeVariants ?? [];
}

/** A real variant's tile scheme: its colours, light and dark swapped when its master is dark. */
function tileScheme(variant: PptxThemeVariant): PptxThemeColorScheme {
	const preset = variant.backgroundStyleIndex
		? backgroundStylePreset(variant.backgroundStyleIndex)
		: undefined;
	const scheme = variant.colorScheme;
	return preset?.dark
		? { ...scheme, dk1: scheme.lt1, lt1: scheme.dk1, dk2: scheme.lt2, lt2: scheme.dk2 }
		: scheme;
}

function buildReal(ctx: RibbonGalleryContext, variants: readonly PptxThemeVariant[]) {
	const current = galleryColorScheme(ctx);
	return {
		id: 'themeVariants' as const,
		labelKey: 'pptx.gallery.themeVariants.title',
		label: 'Variants',
		disabled: false,
		previewable: true,
		inlineCount: 4,
		sections: [
			{
				id: 'variants',
				titleKey: 'pptx.gallery.themeVariants.section.theme',
				title: 'Variants of this theme',
				columns: 4,
				tileWidth: TILE.width,
				tileHeight: TILE.height,
				items: variants.map((variant, index): RibbonGalleryItem => ({
					id: `real-${index}`,
					labelKey: 'pptx.gallery.themeVariants.real',
					labelParams: { number: index + 1, name: variant.name },
					label: `Variant ${index + 1}: ${variant.name}`,
					previewSvg: themeTileSvg(
						tileScheme(variant),
						variant.fontScheme.majorFont?.latin ?? ctx.theme?.fontScheme?.majorFont?.latin,
						TILE,
					),
					applied: colorSchemesEqual(variant.colorScheme, current),
				})),
			},
		],
	};
}

export const THEME_VARIANTS_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const real = realVariants(ctx);
		if (real.length > 0) {
			return buildReal(ctx, real);
		}
		const current = galleryColorScheme(ctx);
		const heading = ctx.theme?.fontScheme?.majorFont?.latin;
		return {
			id: 'themeVariants',
			labelKey: 'pptx.gallery.themeVariants.title',
			label: 'Variants',
			disabled: false,
			previewable: true,
			inlineCount: 4,
			sections: [
				{
					id: 'variants',
					// This theme has no variants of its own: say these four are worked out from its colours.
					titleKey: 'pptx.gallery.themeVariants.section.computed',
					title: 'Computed from the theme colors',
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
		const real = /^real-(\d+)$/u.exec(itemId);
		if (real) {
			const variant = realVariants(ctx)[Number(real[1])];
			return variant ? realVariantResult(variant, ctx) : null;
		}
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

/** A real variant's colours and fonts, its effect scheme and its background style, in order. */
function realVariantResult(variant: PptxThemeVariant, ctx: RibbonGalleryContext) {
	const name = ctx.theme?.name ?? variant.name;
	return {
		kind: 'sequence' as const,
		steps: [
			{
				kind: 'theme' as const,
				colorScheme: { ...variant.colorScheme },
				fontScheme: structuredClone(variant.fontScheme),
				name,
			},
			...(variant.formatSchemeXml
				? [{ kind: 'themeEffects' as const, fmtSchemeXml: variant.formatSchemeXml, name }]
				: []),
			...(variant.backgroundStyleIndex
				? [
						{
							kind: 'backgroundStyle' as const,
							styleIndex: variant.backgroundStyleIndex,
							scope: 'all' as const,
						},
					]
				: []),
		],
	};
}
