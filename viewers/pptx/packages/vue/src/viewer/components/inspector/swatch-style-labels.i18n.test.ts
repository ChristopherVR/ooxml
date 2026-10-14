/**
 * The theme colour grid and the Quick Styles gallery name their swatches through vue-i18n
 * (tooltip and accessible name), not in hard-coded English.
 */
import { mount } from '@vue/test-utils';
import { translationsDe } from 'pptx-viewer-locales';
import { describe, expect, it } from 'vitest';
import { ref } from 'vue';
import { createI18n } from 'vue-i18n';

import { toVueI18nSyntax, translationsEn } from '../../../i18n';
import { ThemeColorMapKey } from '../../composables/theme-color-map-context';
import QuickStylesGallery from './QuickStylesGallery.vue';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.vue';

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

const german = () =>
	createI18n({
		legacy: false,
		locale: 'de',
		fallbackLocale: 'en',
		missingWarn: false,
		fallbackWarn: false,
		messages: { en: translationsEn, de: toVueI18nSyntax(translationsDe) },
	});

describe('swatch and style names follow the locale', () => {
	it('names theme colour swatches with the host dictionary', () => {
		const wrapper = mount(ThemeColorSwatchGrid, {
			global: { plugins: [german()], provide: { [ThemeColorMapKey as symbol]: ref(OFFICE_THEME) } },
		});
		const titles = wrapper.findAll('button').map((button) => button.attributes('title'));
		expect(titles).toContain('Akzent 1, heller 80 %');
		expect(titles.some((title) => title?.includes('Accent'))).toBeFalsy();
		wrapper.unmount();
	});

	it('names Quick Styles presets with the host dictionary', () => {
		const wrapper = mount(QuickStylesGallery, { global: { plugins: [german()] } });
		expect(wrapper.find('button').attributes('aria-label')).toBe('Blaue Füllung');
		wrapper.unmount();
	});
});
