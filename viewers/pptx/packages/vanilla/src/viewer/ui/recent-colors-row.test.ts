import { translationsDe } from 'pptx-viewer-locales';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createRecentColorsRow } from './recent-colors-row';
import { recentSwatches } from './theme-color-swatch-grid.test-support';

describe('createRecentColorsRow', () => {
	it('is hidden until setColors is called with entries', () => {
		const row = createRecentColorsRow(document, createTranslator(), vi.fn());
		expect(row.el.hidden).toBeTruthy();

		row.setColors(['#112233', '#445566']);
		expect(row.el.hidden).toBeFalsy();
		expect(recentSwatches(row.el)).toHaveLength(2);
	});

	it('hides again once setColors is called with an empty list', () => {
		const row = createRecentColorsRow(document, createTranslator(), vi.fn());
		row.setColors(['#112233']);
		expect(row.el.hidden).toBeFalsy();

		row.setColors([]);
		expect(row.el.hidden).toBeTruthy();
	});

	it('is the shared grid with only Recent Colors, carrying the row contract', () => {
		const t = createTranslator();
		const row = createRecentColorsRow(document, t, vi.fn());
		row.setColors(['#112233']);

		expect(row.el.localName).toBe('pptx-ui-color-grid');
		expect(row.el.dataset.testid).toBe('pptx-color-recent');
		expect(row.el.getAttribute('aria-label')).toBe('Recent Colors');
		expect(
			[...row.el.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toStrictEqual(['Recent Colors']);
		const [swatch] = recentSwatches(row.el);
		expect(swatch!.type).toBe('button');
		expect(swatch!.dataset.color).toBe('#112233');
		expect(swatch!.getAttribute('aria-label')).toBe('Recent #112233');
	});

	it('names the row in the viewer language', () => {
		const de = createTranslator('de', { de: translationsDe });
		const row = createRecentColorsRow(document, de, vi.fn());
		row.setColors(['#112233']);
		expect(row.el.getAttribute('aria-label')).toBe('Zuletzt verwendete Farben');
		expect(recentSwatches(row.el)[0]!.getAttribute('aria-label')).toBe(
			'Zuletzt verwendete Farbe #112233',
		);
	});

	it('fires onSelect for the clicked swatch, most-recent first', () => {
		const onSelect = vi.fn();
		const row = createRecentColorsRow(document, createTranslator(), onSelect);
		row.setColors(['#112233', '#445566']);

		recentSwatches(row.el)[1]!.click();

		expect(onSelect).toHaveBeenCalledExactlyOnceWith('#445566');
	});

	it('replaces the row on repeated calls rather than accumulating', () => {
		const row = createRecentColorsRow(document, createTranslator(), vi.fn());
		row.setColors(['#112233', '#445566']);
		row.setColors(['#778899']);

		expect(recentSwatches(row.el)).toHaveLength(1);
	});

	it('setDisabled disables every swatch, including ones added afterward', () => {
		const row = createRecentColorsRow(document, createTranslator(), vi.fn());
		row.setColors(['#112233']);
		row.setDisabled(true);
		expect(recentSwatches(row.el)[0]!.disabled).toBeTruthy();

		row.setColors(['#112233', '#445566']);
		expect(recentSwatches(row.el).every((swatch) => swatch.disabled)).toBeTruthy();
	});
});
