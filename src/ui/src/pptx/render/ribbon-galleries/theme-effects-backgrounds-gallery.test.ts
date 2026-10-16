import type { PptxSlide, ResolvedStyleMatrix, XmlObject } from 'ooxml-core/pptx';
import { THEME_EFFECT_SCHEMES } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n';
import { GALLERY_THEME_PRESETS } from '../../theme/theme-gallery-presets';
import { applyRibbonGalleryItem, buildRibbonGallery } from './gallery-registry';
import type { RibbonGalleryContext } from './gallery-types';

const office = GALLERY_THEME_PRESETS[0]!;

function resolver(calls: Array<{ xml: XmlObject; scheme: string | undefined }>) {
	return (xml: XmlObject, scheme?: string): ResolvedStyleMatrix => {
		calls.push({ xml, scheme });
		const fill = xml['a:fillRef'] as XmlObject;
		const idx = Number(fill['@_idx']);
		return idx >= 1001
			? {
					shapeStyle: {
						fillMode: 'gradient',
						fillColor: '#123456',
						fillGradientStops: [
							{ color: '#112233', position: 0 },
							{ color: '#445566', position: 100 },
						],
						fillGradientAngle: 90,
					},
				}
			: {
					shapeStyle: {
						fillMode: 'solid',
						fillColor: scheme === 'glossy' ? '#AA0000' : '#00AA00',
						strokeColor: '#000000',
						strokeWidth: idx,
						shadowColor: idx === 3 ? '#000000' : undefined,
					},
				};
	};
}

function context(extra: Partial<RibbonGalleryContext> = {}): RibbonGalleryContext {
	return {
		element: null,
		theme: { name: 'Office', colorScheme: office.colorScheme, fontScheme: office.fontScheme },
		...extra,
	};
}

describe('Design > Variants > Effects gallery', () => {
	it('lists every built-in scheme, previewable, and marks the deck scheme applied', () => {
		const descriptor = buildRibbonGallery(
			'themeEffects',
			context({
				theme: {
					name: 'Office',
					colorScheme: office.colorScheme,
					formatScheme: {
						name: 'Office 2007-2010',
						fillStyles: [],
						lineStyles: [],
						effectStyles: [],
						backgroundFillStyles: [],
					},
				},
			}),
		);
		expect(descriptor.disabled).toBe(false);
		expect(descriptor.previewable).toBe(true);
		const items = descriptor.sections.flatMap((section) => section.items);
		expect(items.map((item) => item.id)).toEqual(THEME_EFFECT_SCHEMES.map((scheme) => scheme.id));
		expect(items.filter((item) => item.applied).map((item) => item.id)).toEqual(['office2007']);
		expect(translationsEn[items[0]!.labelKey]).toBe('Effects: {{name}}');
	});

	it('draws each tile with three shapes resolved against that scheme', () => {
		const calls: Array<{ xml: XmlObject; scheme: string | undefined }> = [];
		const descriptor = buildRibbonGallery(
			'themeEffects',
			context({ resolveStyleMatrix: resolver(calls) }),
		);
		const glossy = descriptor.sections[0]!.items.find((item) => item.id === 'glossy')!;
		expect(glossy.previewSvg.match(/<rect /gu)?.length).toBe(3);
		expect(glossy.previewSvg).toContain('#AA0000');
		expect(glossy.previewSvg).toContain('feDropShadow');
		expect(calls.filter((call) => call.scheme === 'glossy').map((call) => call.xml)).toHaveLength(
			3,
		);
	});

	it('draws a neutral tile without a handler', () => {
		const descriptor = buildRibbonGallery('themeEffects', context());
		expect(descriptor.sections[0]!.items[0]!.previewSvg).toContain('<svg');
	});

	it('applies the scheme by id', () => {
		expect(applyRibbonGalleryItem('themeEffects', 'glossy', context())).toEqual({
			kind: 'themeEffects',
			schemeId: 'glossy',
			name: 'Glossy',
		});
		expect(applyRibbonGalleryItem('themeEffects', 'nope', context())).toBeNull();
	});
});

describe('Design > Variants > Background Styles gallery', () => {
	it('is PowerPoint 4 by 3 grid with a context menu per tile and a Format Background footer', () => {
		const descriptor = buildRibbonGallery('themeBackgrounds', context());
		const section = descriptor.sections[0]!;
		expect(descriptor.disabled).toBe(false);
		expect(descriptor.previewable).toBe(true);
		expect(section.columns).toBe(4);
		expect(section.items).toHaveLength(12);
		expect(section.items[0]?.actions?.map((action) => action.id)).toEqual(['all', 'selected']);
		expect(descriptor.footer?.map((action) => action.id)).toEqual(['formatBackground']);
		for (const key of [
			...section.items[0]!.actions!.map((action) => action.labelKey),
			descriptor.footer![0]!.labelKey,
			section.items[0]!.labelKey,
		]) {
			expect(translationsEn[key], key).toBeTruthy();
		}
	});

	it('draws a tile from the background fill style the theme resolves', () => {
		const calls: Array<{ xml: XmlObject; scheme: string | undefined }> = [];
		const items = buildRibbonGallery(
			'themeBackgrounds',
			context({ resolveStyleMatrix: resolver(calls) }),
		).sections[0]!.items;
		expect(items[0]!.previewSvg).toContain('linearGradient');
		// Style 3 is the dark pair of row 1: bg2 swapped to dk2.
		const third = calls.find(
			(call) =>
				(call.xml['a:fillRef'] as XmlObject)['@_idx'] === '1001' &&
				((call.xml['a:fillRef'] as XmlObject)['a:schemeClr'] as XmlObject)['@_val'] === 'dk2',
		);
		expect(third).toBeDefined();
	});

	it('applies to all slides by default, to the selection from the menu, and opens the pane', () => {
		expect(applyRibbonGalleryItem('themeBackgrounds', '5', context())).toEqual({
			kind: 'backgroundStyle',
			styleIndex: 5,
			scope: 'all',
		});
		expect(applyRibbonGalleryItem('themeBackgrounds', '5', context(), 'selected')).toEqual({
			kind: 'backgroundStyle',
			styleIndex: 5,
			scope: 'selected',
		});
		expect(applyRibbonGalleryItem('themeBackgrounds', 'formatBackground', context())).toEqual({
			kind: 'command',
			command: 'formatBackground',
		});
		expect(applyRibbonGalleryItem('themeBackgrounds', '13', context())).toBeNull();
	});

	it('marks the style the active slide uses, the dark columns by its colour map', () => {
		const slide = {
			id: 's',
			backgroundStyleRef: { kind: 'preset', idx: 1002, scheme: 'bg2' },
			clrMapOverride: { bg1: 'dk1', tx1: 'lt1', bg2: 'dk2', tx2: 'lt2' },
		} as unknown as PptxSlide;
		const items = buildRibbonGallery('themeBackgrounds', context({ activeSlide: slide }))
			.sections[0]!.items;
		expect(items.filter((item) => item.applied).map((item) => item.id)).toEqual(['7']);
	});
});
