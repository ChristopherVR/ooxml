/**
 * Swatch and style names reach the DOM as tooltips and accessible names, so they must come from
 * the binding's dictionary like every other label: a host translating the viewer into a language
 * the repo does not ship (here a Korean map) must see its own words, and English must stay the
 * fallback for any key the host leaves out.
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { createTranslator } from 'ooxml-core/i18n';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../i18n/translations-en';
import { ACTION_BUTTON_PRESETS, actionButtonLabelKey } from './action-buttons';
import { OFFICE_COLOR_SWATCHES, officeColorSwatchLabel } from './color-swatches';
import { galleryItemLabel } from './ribbon-galleries/gallery-view';
import { buildRibbonGallery } from './ribbon-galleries/gallery-registry';
import type { RibbonGalleryItem } from './ribbon-galleries/gallery-types';
import { SHAPE_QUICK_STYLES, shapeQuickStyleLabel } from './shape-quick-styles';
import {
	buildThemeColorSwatchGrid,
	THEME_COLOR_SWATCH_COLUMN_LABEL_KEYS,
	THEME_COLOR_SWATCH_VARIANT_LABEL_KEYS,
} from './theme-color-swatches';

const KO: Record<string, string> = {
	'pptx.themeColor.accent1': '강조 1',
	'pptx.themeColor.background1': '배경 1',
	'pptx.themeColor.swatchLighter': '{{color}}, {{percent}}% 더 밝게',
	'pptx.themeColor.swatchDarker': '{{color}}, {{percent}}% 더 어둡게',
	'pptx.colorPicker.standard.red': '빨강',
	'pptx.shape.quickStylePreset.blueFill': '파랑 채우기',
	'pptx.gallery.shapeStyles.coloredFill': '색 채우기 - {{color}}',
	'pptx.gallery.themeColors.name.blue-warm': '따뜻한 파랑',
	'pptx.gallery.themeFonts.name.office-2013-2022': '오피스 2013 - 2022',
	'pptx.gallery.themeFonts.scheme': '{{name}}: {{major}} / {{minor}}',
	'pptx.insert.actionButton.backPrevious': '뒤로 또는 이전',
};

/** A host translator: its own map first, then the shipped English dictionary. */
const ko = createTranslator(KO, [translationsEn], { placeholder: 'double' });
/** A host translator whose map has none of these keys (answers a miss with the key). */
const empty = createTranslator({}, [], { placeholder: 'double' });

const OFFICE_THEME = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	bg1: '#FFFFFF',
	tx1: '#000000',
	bg2: '#E7E6E6',
	tx2: '#44546A',
};

const textShape = {
	id: 's1',
	type: 'shape',
	x: 0,
	y: 0,
	width: 200,
	height: 100,
	shapeType: 'rect',
	shapeStyle: {},
	text: 'Text',
	textSegments: [{ text: 'Text', style: {} }],
} as unknown as PptxElement;

function galleryItems(id: 'shapeStyles' | 'themeColors' | 'themeFonts'): RibbonGalleryItem[] {
	return buildRibbonGallery(id, {
		element: textShape,
		themeColorMap: OFFICE_THEME,
	}).sections.flatMap((section) => section.items);
}

describe('theme colour swatch names', () => {
	it('names the theme grid in the host language', () => {
		const accent1 = buildThemeColorSwatchGrid(OFFICE_THEME, ko).find(
			(column) => column.scheme === 'accent1',
		)!;
		expect(accent1.base.label).toBe('강조 1');
		expect(accent1.variants.map((swatch) => swatch.label)).toStrictEqual([
			'강조 1, 80% 더 밝게',
			'강조 1, 60% 더 밝게',
			'강조 1, 40% 더 밝게',
			'강조 1, 25% 더 어둡게',
			'강조 1, 50% 더 어둡게',
		]);
		const background1 = buildThemeColorSwatchGrid(OFFICE_THEME, ko)[0]!;
		expect(background1.variants[0]!.label).toBe('배경 1, 5% 더 어둡게');
	});

	it('keeps English for a host dictionary without the keys, and without a translator', () => {
		for (const translate of [empty, undefined]) {
			const accent1 = buildThemeColorSwatchGrid(OFFICE_THEME, translate).find(
				(column) => column.scheme === 'accent1',
			)!;
			expect(accent1.variants[0]!.label).toBe('Accent 1, Lighter 80%');
			expect(accent1.variants[3]!.label).toBe('Accent 1, Darker 25%');
		}
	});

	it('ships every key the grid reads in the English dictionary, with the same English text', () => {
		const keys = [
			...Object.values(THEME_COLOR_SWATCH_COLUMN_LABEL_KEYS),
			...Object.values(THEME_COLOR_SWATCH_VARIANT_LABEL_KEYS),
		];
		for (const key of keys) {
			expect(translationsEn, key).toHaveProperty([key]);
		}
		const en = createTranslator(translationsEn, [], { placeholder: 'double' });
		expect(buildThemeColorSwatchGrid(OFFICE_THEME, en)).toStrictEqual(
			buildThemeColorSwatchGrid(OFFICE_THEME),
		);
	});
});

describe('standard colour and quick style names', () => {
	it('translates a standard swatch and falls back to English', () => {
		const red = OFFICE_COLOR_SWATCHES.find((swatch) => swatch.hex === '#ff0000')!;
		expect(officeColorSwatchLabel(red, ko)).toBe('빨강');
		expect(officeColorSwatchLabel(red, empty)).toBe('Red');
		for (const swatch of OFFICE_COLOR_SWATCHES) {
			expect(translationsEn[swatch.labelKey as keyof typeof translationsEn]).toBe(swatch.label);
		}
	});

	it('translates a quick style preset and falls back to English', () => {
		const [blueFill] = SHAPE_QUICK_STYLES;
		expect(shapeQuickStyleLabel(blueFill!, ko)).toBe('파랑 채우기');
		expect(shapeQuickStyleLabel(blueFill!, empty)).toBe('Blue Fill');
		for (const preset of SHAPE_QUICK_STYLES) {
			expect(translationsEn[preset.labelKey as keyof typeof translationsEn]).toBe(preset.name);
		}
	});

	it('translates the Insert > Action button names', () => {
		const back = ACTION_BUTTON_PRESETS[0]!;
		expect(ko(actionButtonLabelKey(back.shapeType))).toBe('뒤로 또는 이전');
		for (const preset of ACTION_BUTTON_PRESETS) {
			const key = actionButtonLabelKey(preset.shapeType);
			expect(translationsEn[key as keyof typeof translationsEn], key).toBe(preset.label);
		}
	});
});

describe('gallery tile names built from a translatable name', () => {
	it('translates the colour inside a Shape Styles tile name', () => {
		const tile = galleryItems('shapeStyles').find(
			(item) => item.labelKey === 'pptx.gallery.shapeStyles.coloredFill',
		)!;
		expect(galleryItemLabel(tile, ko)).toMatch(/^색 채우기 - (강조 1|Dark 1)$/u);
		const accent1 = galleryItems('shapeStyles').find(
			(item) =>
				item.labelKey === 'pptx.gallery.shapeStyles.coloredFill' &&
				item.labelParams?.color === 'Accent 1',
		)!;
		expect(galleryItemLabel(accent1, ko)).toBe('색 채우기 - 강조 1');
	});

	it('translates Design > Variants colour and font scheme names, keeping font names', () => {
		const colors = galleryItems('themeColors');
		const blueWarm = colors.find((item) => item.label === 'Blue Warm')!;
		expect(galleryItemLabel(blueWarm, ko)).toBe('따뜻한 파랑');
		const fonts = galleryItems('themeFonts');
		const office = fonts.find((item) => item.labelParams?.name === 'Office 2013 - 2022')!;
		expect(galleryItemLabel(office, ko)).toBe('오피스 2013 - 2022: Calibri Light / Calibri');
		const calibri = fonts.find((item) => item.labelParams?.name === 'Calibri')!;
		expect(calibri.labelParamKeys).toBeUndefined();
		for (const item of [...colors, ...fonts]) {
			for (const { key } of Object.values(item.labelParamKeys ?? {})) {
				expect(translationsEn, key).toHaveProperty([key]);
			}
		}
	});
});
