// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { localizeRibbon, ribbonLocale } from './ribbon-localize';
import {
	VISIO_LOCALES,
	VISIO_RIBBON_STRINGS,
	normalizeVisioLocale,
	translateRibbonText,
} from './ribbon-locale';

const english = Object.keys(VISIO_RIBBON_STRINGS.en);
const placeholders = (text: string) => [...text.matchAll(/\{\w+\}/g)].map((m) => m[0]).sort();

describe('Visio ribbon locales', () => {
	it('has a dictionary for each supported locale', () => {
		expect([...VISIO_LOCALES]).toEqual(['en', 'de', 'es', 'fr', 'zh-CN']);
		expect(Object.keys(VISIO_RIBBON_STRINGS).sort()).toEqual([...VISIO_LOCALES].sort());
	});

	for (const locale of VISIO_LOCALES) {
		describe(locale, () => {
			it('defines every English key, with no extras and no empty strings', () => {
				expect(Object.keys(VISIO_RIBBON_STRINGS[locale]).sort()).toEqual([...english].sort());
				for (const key of english)
					expect(VISIO_RIBBON_STRINGS[locale][key]?.trim(), key).toBeTruthy();
			});

			it('keeps the placeholders of every English template', () => {
				for (const key of english)
					expect(placeholders(VISIO_RIBBON_STRINGS[locale][key] ?? ''), `${locale}:${key}`).toEqual(
						placeholders(key),
					);
			});

			if (locale !== 'en')
				it('translates the long sentences', () => {
					const same = english.filter(
						(key) => key.split(' ').length > 3 && VISIO_RIBBON_STRINGS[locale][key] === key,
					);
					expect(same).toEqual([]);
				});
		});
	}

	it('normalizes locale tags', () => {
		expect(normalizeVisioLocale('de-AT')).toBe('de');
		expect(normalizeVisioLocale('zh-Hans')).toBe('zh-CN');
		expect(normalizeVisioLocale('es_MX')).toBe('es');
		expect(normalizeVisioLocale('ja')).toBe('en');
		expect(normalizeVisioLocale(undefined)).toBe('en');
	});

	it('translates tooltips, templates and unknown text', () => {
		expect(translateRibbonText('de', 'Bold (Ctrl+B)')).toBe('Fett (Ctrl+B)');
		expect(translateRibbonText('fr', 'Pattern 7')).toBe('Motif 7');
		expect(translateRibbonText('es', 'Offset: Bottom Right')).toBe(
			'Desplazamiento: Abajo a la derecha',
		);
		expect(
			translateRibbonText('de', 'Pointer Tool: not available yet. Ink strokes are not supported.'),
		).toBe('Zeigerwerkzeug: noch nicht verfügbar. Freihandstriche werden nicht unterstützt.');
		expect(translateRibbonText('zh-CN', 'A custom shape')).toBe('A custom shape');
		expect(translateRibbonText('en', 'Bold')).toBe('Bold');
	});
});

describe('localizeRibbon', () => {
	const tabLabels = (ribbon: HTMLElement) =>
		[...ribbon.children].map((panel) => (panel as HTMLElement).dataset.label);

	it('translates tabs, labels and tooltips, and back again from the English originals', () => {
		const ribbon = createRibbon(document);
		expect(localizeRibbon(ribbon, 'de-DE')).toBe('de');
		expect(tabLabels(ribbon).filter(Boolean).slice(0, 3)).toEqual(['Start', 'Einfügen', 'Entwurf']);
		const bold = ribbon.querySelector('[command="bold"]')!;
		expect(bold.getAttribute('label')).toBe('Fett');
		expect(bold.getAttribute('title')).toContain('Fett');
		expect(ribbon.getAttribute('file-label')).toBe('Datei');
		expect(ribbonLocale(ribbon)).toBe('de');
		localizeRibbon(ribbon, 'zh');
		expect(bold.getAttribute('label')).toBe('加粗');
		localizeRibbon(ribbon, 'en');
		expect(bold.getAttribute('label')).toBe('Bold');
		expect(tabLabels(ribbon).filter(Boolean)[0]).toBe('Home');
		expect(ribbon.getAttribute('file-label')).toBe('File');
	});

	it('leaves a host add-in tab alone', () => {
		const ribbon = createRibbon(document);
		const addIn = document.createElement('div');
		addIn.dataset.ribbonTab = 'host';
		addIn.dataset.label = 'Design';
		addIn.dataset.addIn = '';
		ribbon.append(addIn);
		localizeRibbon(ribbon, 'de');
		expect(addIn.dataset.label).toBe('Design');
	});
});
