// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { VISIO_SHAPES_STRINGS } from './locales/shapes';
import { VISIO_STENCIL_STRINGS } from './locales/stencils';
import { localizeRibbon, ribbonLocale } from './ribbon-localize';
import {
	VISIO_LOCALES,
	VISIO_RIBBON_STRINGS,
	normalizeVisioLocale,
	translateRibbonText,
} from './ribbon-locale';

const english = Object.keys(VISIO_RIBBON_STRINGS.en);
const placeholders = (text: string) => [...text.matchAll(/\{\w+\}/g)].map((m) => m[0]).sort();

it('keeps the Shapes pane and stencil dictionaries complete in every locale', () => {
	for (const strings of [VISIO_SHAPES_STRINGS, VISIO_STENCIL_STRINGS]) {
		const englishEntries = Object.entries(strings.en);
		for (const locale of VISIO_LOCALES) {
			const dictionary: Readonly<Record<string, string>> = strings[locale];
			expect(Object.keys(dictionary).sort()).toEqual(englishEntries.map(([key]) => key).sort());
			for (const [key, value] of englishEntries) {
				const translation = dictionary[key] ?? '';
				expect(translation.trim(), `${locale}:${key}`).toBeTruthy();
				expect(placeholders(translation), `${locale}:${key}`).toEqual(placeholders(value));
			}
		}
	}
});

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
		const first = 'Comment on the current page or selected shape.';
		const second = 'An approximate picture with compatibility notes.';
		expect(translateRibbonText('de', `${first}  ${second}`)).toBe(
			`${VISIO_RIBBON_STRINGS.de[first]} ${VISIO_RIBBON_STRINGS.de[second]}`,
		);
		expect(translateRibbonText('de', `${first}${second}`)).toBe(`${first}${second}`);
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

	it('stays fast on long crafted text (code scanning: polynomial patterns)', () => {
		const started = performance.now();
		translateRibbonText('de', `${'-'.repeat(20000)}: x`);
		translateRibbonText('de', `a${'		'.repeat(20000)}`);
		translateRibbonText('de', `${'a. '.repeat(5000)}b`);
		translateRibbonText('de', `${'!'.repeat(20000)}`);
		expect(performance.now() - started).toBeLessThan(500);
		expect(translateRibbonText('de', 'Bold: Bold')).toBe('Fett: Fett');
	});
});

describe('load refusals', () => {
	it('tells the user in their language how many shapes a page has that is too large to open', () => {
		const refusal =
			'The page has 12000 shapes, more than this viewer can open (it reads at most 1250000 XML nodes per part).';
		expect(translateRibbonText('de', refusal)).toBe(
			'Die Seite enthält 12000 Formen, mehr als dieser Viewer öffnen kann (er liest höchstens 1250000 XML-Knoten pro Teil).',
		);
		for (const locale of VISIO_LOCALES.filter((locale) => locale !== 'en')) {
			for (const message of [
				refusal,
				'A part of this drawing has more than 1250000 XML nodes, more than this viewer can open.',
				'This drawing has more than 2000000 XML nodes in all, more than this viewer can open.',
				'This drawing would need more than 600 MB of memory to open, more than this viewer allows.',
				'This 48 KB drawing would need more than 189 MB of memory to open, far more than drawings of its size need.',
			]) {
				const translated = translateRibbonText(locale, message);
				expect(translated, locale).not.toBe(message);
				for (const number of message.match(/\d+/g)!) expect(translated).toContain(number);
			}
		}
	});
});

describe('edit refusals', () => {
	it('tells the user in their language why a drawing cannot be edited', () => {
		expect(
			translateRibbonText(
				'de',
				'Edit rejected: This drawing is too large to edit: an edit follows at most 100000 ShapeSheet cells.',
			),
		).toBe(
			'Bearbeitung abgelehnt: Diese Zeichnung ist zu groß zum Bearbeiten: Eine Bearbeitung folgt höchstens 100000 ShapeSheet-Zellen.',
		);
		for (const locale of VISIO_LOCALES.filter((locale) => locale !== 'en')) {
			for (const message of [
				'Edit rejected: This drawing is too large to edit: an edit follows at most 100000 ShapeSheet cells.',
				'Edit rejected: This drawing would need more than 600 MB of memory to edit, more than this viewer allows.',
				'Edit rejected: This 48 KB drawing would need more than 756 MB of memory to edit, far more than drawings of its size need.',
				'This drawing is too large to edit here: its pages hold about 260000 ShapeSheet cells, more than the 100000 an edit can follow. It opens read-only.',
				'This drawing is too large to edit here: an edit would need more than 480 MB of memory. It opens read-only.',
			]) {
				const translated = translateRibbonText(locale, message);
				expect(translated, locale).not.toMatch(/Edit rejected|too large|would need|read-only/);
				for (const number of message.match(/\d+/g)!) expect(translated).toContain(number);
			}
		}
	});
});
