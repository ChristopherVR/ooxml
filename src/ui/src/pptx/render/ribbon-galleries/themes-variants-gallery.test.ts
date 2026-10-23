import type { PptxThemeVariant } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n';
import { GALLERY_THEME_PRESETS } from '../../theme/theme-gallery-presets';
import { FIXED_TAB_GALLERIES } from './gallery-placements';
import { applyRibbonGalleryItem, buildRibbonGallery } from './gallery-registry';
import { inlineGalleryItems } from './gallery-view';
import { themeVariants } from './theme-variants-gallery';

const office = GALLERY_THEME_PRESETS[0]!;
const facet = GALLERY_THEME_PRESETS[1]!;
const ctx = (preset = office, name = preset.name) => ({
	element: null,
	theme: { name, colorScheme: preset.colorScheme, fontScheme: preset.fontScheme },
});

describe('Design > Themes gallery', () => {
	it('offers the current and every built-in theme, three inline, and marks the deck theme', () => {
		const descriptor = buildRibbonGallery('themes', ctx(facet));
		const items = descriptor.sections.flatMap((section) => section.items);
		expect(items.map((item) => item.id)).toEqual([
			'current',
			...GALLERY_THEME_PRESETS.map((preset) => preset.id),
		]);
		expect(inlineGalleryItems(descriptor)).toHaveLength(3);
		expect(items.filter((item) => item.applied).map((item) => item.id)).toEqual(['facet']);
		expect(items[0]?.previewSvg).toContain('<svg');
	});

	it('keeps every tile in place when a built-in theme replaces a custom one', () => {
		const custom = {
			element: null,
			theme: { name: 'Mine', colorScheme: { ...office.colorScheme, accent1: '#123456' } },
		};
		const ids = (context: Parameters<typeof buildRibbonGallery>[1]) =>
			buildRibbonGallery('themes', context).sections[0]!.items.map((item) => item.id);
		// The picked tile stays under the pointer, so it does not start previewing another theme.
		expect(ids(ctx(facet))).toEqual(ids(custom));
	});

	it('heads the list with the current theme when it is not a built-in one', () => {
		const custom = {
			element: null,
			theme: { name: 'Mine', colorScheme: { ...office.colorScheme, accent1: '#123456' } },
		};
		const items = buildRibbonGallery('themes', custom).sections[0]!.items;
		expect(items[0]).toMatchObject({ id: 'current', applied: true });
	});

	it('applies the whole theme: colours, fonts and name', () => {
		const result = applyRibbonGalleryItem('themes', 'facet', ctx());
		expect(result).toEqual({
			kind: 'theme',
			colorScheme: facet.colorScheme,
			fontScheme: facet.fontScheme,
			name: 'Facet',
		});
		expect(applyRibbonGalleryItem('themes', 'nope', ctx())).toBeNull();
	});
});

describe('Design > Variants gallery', () => {
	it('derives four variants from the theme colours', () => {
		const [original, inverted, shiftTwo, shiftFour] = themeVariants(office.colorScheme);
		expect(original?.scheme).toEqual(office.colorScheme);
		expect(inverted?.scheme).toMatchObject({
			dk1: office.colorScheme.lt1,
			lt1: office.colorScheme.dk1,
			dk2: office.colorScheme.lt2,
			lt2: office.colorScheme.dk2,
			accent1: office.colorScheme.accent1,
		});
		expect(shiftTwo?.scheme.accent1).toBe(office.colorScheme.accent3);
		expect(shiftTwo?.scheme.accent5).toBe(office.colorScheme.accent1);
		expect(shiftFour?.scheme.accent1).toBe(office.colorScheme.accent5);
	});

	it('writes a colour scheme only and keeps the theme name', () => {
		const result = applyRibbonGalleryItem('themeVariants', 'inverted', ctx());
		expect(result).toMatchObject({ kind: 'themeColorScheme', name: 'Office' });
		expect(result && 'fontScheme' in result).toBe(false);
	});

	it('keeps the same four tiles after a variant is applied', () => {
		const applied = applyRibbonGalleryItem('themeVariants', 'shiftTwo', ctx());
		if (applied?.kind !== 'themeColorScheme') {
			throw new Error('expected a colour scheme');
		}
		const after = { element: null, theme: { name: 'Office', colorScheme: applied.colorScheme } };
		const items = buildRibbonGallery('themeVariants', after).sections[0]!.items;
		expect(items.map((item) => item.applied)).toEqual([false, false, true, false]);
	});

	it('recognises the theme PowerPoint names "Office Theme" after a variant changed its colours', () => {
		const applied = applyRibbonGalleryItem(
			'themeVariants',
			'inverted',
			ctx(office, 'Office Theme'),
		);
		if (applied?.kind !== 'themeColorScheme') {
			throw new Error('expected a colour scheme');
		}
		const after = {
			element: null,
			theme: { name: 'Office Theme', colorScheme: applied.colorScheme },
		};
		const items = buildRibbonGallery('themeVariants', after).sections[0]!.items;
		expect(items.map((item) => item.applied)).toEqual([false, true, false, false]);
	});

	it('labels the computed variants honestly and previews them on hover', () => {
		const descriptor = buildRibbonGallery('themeVariants', ctx());
		expect(descriptor.previewable).toBe(true);
		expect(descriptor.sections[0]?.titleKey).toBe('pptx.gallery.themeVariants.section.computed');
		expect(translationsEn[descriptor.sections[0]!.titleKey!]).toBe(descriptor.sections[0]!.title);
	});

	it('offers the real variants a package carries instead of computed ones', () => {
		const variants: PptxThemeVariant[] = [
			{ id: '{A}', name: 'Facet', colorScheme: facet.colorScheme, fontScheme: facet.fontScheme },
			{
				id: '{B}',
				name: 'Facet',
				colorScheme: { ...facet.colorScheme, accent1: '#CC0000' },
				fontScheme: facet.fontScheme,
				formatSchemeXml: '<a:fmtScheme name="Glossy"></a:fmtScheme>',
				backgroundStyleIndex: 8,
			},
		];
		const real = { ...ctx(facet), themeVariants: variants };
		const descriptor = buildRibbonGallery('themeVariants', real);
		expect(descriptor.sections[0]?.titleKey).toBe('pptx.gallery.themeVariants.section.theme');
		expect(descriptor.sections[0]?.items.map((item) => [item.id, item.applied])).toEqual([
			['real-0', true],
			['real-1', false],
		]);
		expect(descriptor.sections[0]?.items[1]?.label).toBe('Variant 2: Facet');
		expect(applyRibbonGalleryItem('themeVariants', 'real-0', real)).toMatchObject({
			kind: 'sequence',
			steps: [{ kind: 'theme', name: 'Facet' }],
		});
		const second = applyRibbonGalleryItem('themeVariants', 'real-1', real);
		expect(second).toMatchObject({
			kind: 'sequence',
			steps: [
				{ kind: 'theme', colorScheme: { accent1: '#CC0000' } },
				{ kind: 'themeEffects', fmtSchemeXml: '<a:fmtScheme name="Glossy"></a:fmtScheme>' },
				{ kind: 'backgroundStyle', styleIndex: 8, scope: 'all' },
			],
		});
		expect(applyRibbonGalleryItem('themeVariants', 'real-9', real)).toBeNull();
	});

	it('is placed in the Design tab next to Colors and Fonts', () => {
		expect(
			FIXED_TAB_GALLERIES.filter((placement) => placement.control.startsWith('design.')).map(
				(placement) => [placement.control, placement.mode],
			),
		).toEqual([
			['design.themes.gallery', 'inline'],
			['design.variants.gallery', 'inline'],
			['design.variants.colors', 'dropdown'],
			['design.variants.fonts', 'dropdown'],
			['design.variants.effects', 'dropdown'],
			['design.variants.backgroundStyles', 'dropdown'],
		]);
	});
});
