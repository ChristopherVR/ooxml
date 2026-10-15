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
	it('offers every built-in theme, three of them inline, and marks the deck theme applied', () => {
		const descriptor = buildRibbonGallery('themes', ctx(facet));
		const items = descriptor.sections.flatMap((section) => section.items);
		expect(items.map((item) => item.id)).toEqual(GALLERY_THEME_PRESETS.map((preset) => preset.id));
		expect(inlineGalleryItems(descriptor)).toHaveLength(3);
		expect(items.filter((item) => item.applied).map((item) => item.id)).toEqual(['facet']);
		expect(items[0]?.previewSvg).toContain('<svg');
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

	it('lists Effects and Background Styles disabled, with the reason', () => {
		for (const id of ['themeEffects', 'themeBackgrounds'] as const) {
			const descriptor = buildRibbonGallery(id, ctx());
			expect(descriptor.disabled).toBe(true);
			expect(descriptor.command?.hintKey).toMatch(/unavailable$/u);
			expect(translationsEn[descriptor.command!.hintKey!]).toBe(descriptor.command!.hint);
			expect(applyRibbonGalleryItem(id, 'run', ctx())).toBeNull();
		}
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
