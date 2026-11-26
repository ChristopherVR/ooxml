import { describe, expect, it } from 'vitest';
import { mergeMessageModules, normalizeLocale } from '../../i18n/index';
import { FUNCTION_CATALOG } from '../formula/functions/registry';
import { COMMAND_TABLES } from './commands/i18n/index';
import { DRAWING_FORMAT_STRINGS } from './commands/i18n/drawing-format';
import {
	EDITOR_LOCALES,
	STRINGS,
	loadEditorStrings,
	normalizeEditorLocale,
	translate,
} from './localization';
import { mergeStringModules } from './locales/merge';

describe('xlsx localization over the shared i18n area', () => {
	it('normalises like the shared helper', () => {
		for (const tag of ['de-AT', 'zh_Hans_CN', 'zh-TW', 'pt', 'FR', null]) {
			expect(normalizeEditorLocale(tag)).toBe(normalizeLocale(tag, EDITOR_LOCALES, 'en'));
		}
	});
	it('merges like the shared helper and translates with English fallback', () => {
		const modules = { './a.ts': { a: { x: 'X' } } };
		expect(mergeStringModules(modules)).toEqual(mergeMessageModules(modules));
		expect(translate('de', 'definitely missing {n}', { n: 1 })).toBe('definitely missing 1');
		const key = Object.keys(STRINGS.en)[0] ?? '';
		expect(translate('fr', key)).toBe(STRINGS.fr[key] ?? STRINGS.en[key]);
	});
	it('loads only the selected language, deduplicates concurrent loads and retains English fallback', async () => {
		expect(translate('fr', 'Bold')).toBe('Bold');
		const french = loadEditorStrings(normalizeEditorLocale('fr-CA'));
		expect(loadEditorStrings('fr')).toBe(french);
		expect((await french).Bold).toBe('Gras');
		expect(translate('fr', 'Bold')).toBe('Gras');
		expect(await loadEditorStrings('fr')).toBe(await french);
		expect(translate('de', 'Bold')).toBe('Bold');
		expect(translate('fr', 'Unregistered {value}', { value: 5 })).toBe('Unregistered 5');
	});
	it('has four non-empty translations per command row, keeping placeholders and no em-dash', () => {
		const problems: string[] = [];
		const placeholders = (text: string) => (text.match(/\{\w+\}/g) ?? []).sort().join();
		for (const table of [...COMMAND_TABLES, DRAWING_FORMAT_STRINGS])
			for (const [key, texts] of Object.entries(table)) {
				if (texts.length !== 4) problems.push(`${key} needs four translations`);
				for (const [index, text] of texts.entries()) {
					if (!text.trim()) problems.push(`${key} [${index}] empty`);
					if (text.includes('\u2014')) problems.push(`${key} [${index}] em-dash`);
					if (placeholders(text) !== placeholders(key))
						problems.push(`${key} [${index}] placeholders`);
				}
				if (key.includes('\u2014')) problems.push(`${key} em-dash`);
			}
		expect(problems).toEqual([]);
	});
	it('provides non-empty translations for all FUNCTION_CATALOG descriptions across all locales without em-dashes', async () => {
		const nonEnglishLocales = EDITOR_LOCALES.filter((loc) => loc !== 'en');
		const tables = await Promise.all(
			nonEnglishLocales.map(async (locale) => ({
				locale,
				table: await loadEditorStrings(locale),
			})),
		);
		expect(FUNCTION_CATALOG.length).toBeGreaterThan(500);

		for (const fn of FUNCTION_CATALOG) {
			const desc = fn.description;
			expect(desc).toBeTruthy();
			expect(STRINGS.en[desc]).toBe(desc);

			for (const { locale, table } of tables) {
				const translated = table[desc];
				expect(translated, `Missing ${locale} translation for ${fn.name}: "${desc}"`).toBeTruthy();
				expect(translated).not.toContain('\u2014');
			}
		}
	});
});
