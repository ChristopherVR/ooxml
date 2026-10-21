/**
 * The theme colour grid and the Quick Styles gallery name their swatches in the viewer's locale
 * (tooltip and accessible name), not in hard-coded English.
 */
import { translationsDe } from 'pptx-viewer-locales';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createQuickStylesGallery } from './inspector/quick-styles-gallery';
import { createThemeColorSwatchGrid } from './theme-color-swatch-grid';
import { themeSwatches } from './theme-color-swatch-grid.test-support';

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

const de = createTranslator('de', { de: translationsDe });

describe('swatch and style names follow the locale', () => {
	it('names theme colour swatches with the host dictionary', () => {
		const grid = createThemeColorSwatchGrid(document, de, vi.fn());
		grid.setThemeColorMap(OFFICE_THEME);
		const titles = themeSwatches(grid.el).map((button) => button.title);
		expect(titles).toContain('Akzent 1');
		expect(titles).toContain('Akzent 1, heller 80 %');
		expect(titles).toContain('Akzent 1, dunkler 25 %');
		expect(titles.some((title) => title.includes('Accent'))).toBeFalsy();
		// The heading and the grid's accessible name are translated too.
		expect(grid.el.shadowRoot!.querySelector('.heading')!.textContent).toBe('Designfarben');
		expect(grid.el.getAttribute('aria-label')).toBe('Designfarben');
	});

	it('names Quick Styles presets with the host dictionary', () => {
		const gallery = createQuickStylesGallery(
			document,
			de,
			() => document.createElement('section'),
			{ setShapeStyle: vi.fn() },
		);
		const first = gallery.el.querySelector<HTMLButtonElement>('button')!;
		expect(first.title).toBe('Blaue Füllung');
		expect(first.getAttribute('aria-label')).toBe('Blaue Füllung');
	});
});
