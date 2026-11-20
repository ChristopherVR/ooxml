// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { labelNameMismatches } from '../test-support/label-in-name';
import { EDITOR_LOCALES, findLocalizedControl } from './localization';
import { createRibbon } from './ribbon';

describe('Word ribbon names (WCAG 2.5.3)', () => {
	for (const locale of EDITOR_LOCALES)
		it(`names every captioned command with its caption in ${locale}`, () => {
			const ribbon = createRibbon(locale);
			document.body.append(ribbon);
			try {
				expect(labelNameMismatches(ribbon)).toEqual([]);
			} finally {
				ribbon.remove();
			}
		});
});

describe('Word ribbon captions stay as Word shows them', () => {
	const shown = (ribbon: HTMLElement, id: string) => {
		const control = findLocalizedControl<HTMLButtonElement>(ribbon, id)!;
		return {
			caption: control.textContent?.trim(),
			name: control.getAttribute('aria-label'),
			title: control.title,
		};
	};
	it.each([
		['en', 'Insert picture', 'Pictures', 'Insert pictures', 'Insert picture'],
		['de', 'Insert picture', 'Bilder', 'Bilder einfügen', 'Bild einfügen'],
		['fr', 'Insert picture', 'Images', 'Insérer des images', 'Insérer une image'],
		['zh-CN', 'Select all', '选择', '全部选择', '全选 (Ctrl+A)'],
		[
			'en',
			'Different odd and even pages',
			'Different odd & even pages',
			'Different odd & even pages',
			'Different odd and even pages',
		],
		[
			'de',
			'Table properties',
			'Eigenschaften',
			'Eigenschaften der Tabelle',
			'Tabelleneigenschaften',
		],
	])('%s %s: caption, name and tooltip', (locale, id, caption, name, title) => {
		const ribbon = createRibbon(locale);
		document.body.append(ribbon);
		try {
			expect(shown(ribbon, id)).toEqual({ caption, name, title });
		} finally {
			ribbon.remove();
		}
	});
});
