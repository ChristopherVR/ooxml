import { describe, expect, it } from 'vitest';
import { mergeMessageModules, normalizeLocale } from '../../i18n/index';
import { EDITOR_LOCALES, STRINGS, normalizeEditorLocale, translate } from './localization';
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
});
