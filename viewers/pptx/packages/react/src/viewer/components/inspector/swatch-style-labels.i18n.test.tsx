// @vitest-environment happy-dom
/**
 * The theme colour grid and the Quick Styles gallery name their swatches through the host's
 * i18next dictionary (tooltip and accessible name), not in hard-coded English.
 */
import { createInstance } from 'i18next';
import { translationsDe } from 'pptx-viewer-locales';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { renderToStaticMarkup } from 'react-dom/server';
import { I18nextProvider, initReactI18next } from 'react-i18next';
import { describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../i18n';
import { QuickStylesGallery } from './QuickStylesGallery';
import { ThemeColorMapProvider } from './ThemeColorMapContext';
import { ThemeColorSwatchGrid } from './ThemeColorSwatchGrid';

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

async function german() {
	const instance = createInstance();
	await instance.use(initReactI18next).init({
		lng: 'de',
		fallbackLng: 'en',
		resources: { en: { translation: translationsEn }, de: { translation: translationsDe } },
		interpolation: { escapeValue: false },
	});
	return instance;
}

function render(node: React.ReactElement, i18n: Awaited<ReturnType<typeof german>>) {
	const container = document.createElement('div');
	container.innerHTML = renderToStaticMarkup(<I18nextProvider i18n={i18n}>{node}</I18nextProvider>);
	return container;
}

describe('swatch and style names follow the locale', () => {
	it('names theme colour swatches with the host dictionary', async () => {
		// The shared grid draws its swatches in its shadow root, so mount it instead of static markup.
		const i18n = await german();
		const container = document.body.appendChild(document.createElement('div'));
		const root = createRoot(container);
		act(() =>
			root.render(
				<I18nextProvider i18n={i18n}>
					<ThemeColorMapProvider value={OFFICE_THEME}>
						<ThemeColorSwatchGrid onPick={vi.fn()} />
					</ThemeColorMapProvider>
				</I18nextProvider>,
			),
		);
		const shadow = container.querySelector('pptx-ui-color-grid')!.shadowRoot!;
		const titles = [...shadow.querySelectorAll('button')].map((button) => button.title);
		expect(titles).toContain('Akzent 1, heller 80 %');
		expect(titles.some((title) => title.includes('Accent'))).toBeFalsy();
		expect(shadow.querySelector('.heading')!.textContent).toBe('Designfarben');
		act(() => root.unmount());
		container.remove();
	});

	it('names Quick Styles presets with the host dictionary', async () => {
		const container = render(<QuickStylesGallery onUpdateShapeStyle={vi.fn()} />, await german());
		expect(container.querySelector('button')!.getAttribute('aria-label')).toBe('Blaue Füllung');
	});
});
