import { missingMessageKeys } from 'ooxml-core/i18n';
import { TEAMS_LOCALES, TEAMS_MESSAGES, normalizeTeamsLocale, teamsI18n } from './translator';

const PLACEHOLDER = /\{(\w+)\}/gu;
const placeholders = (text: string): string[] =>
	[...text.matchAll(PLACEHOLDER)].map((m) => m[1]!).sort();

const english = TEAMS_MESSAGES.en;
const others = TEAMS_LOCALES.filter((locale) => locale !== 'en');

describe('Teams dictionaries', () => {
	it('ships English plus de, es, fr and zh-CN', () => {
		expect([...TEAMS_LOCALES].sort()).toEqual(['de', 'en', 'es', 'fr', 'zh-CN']);
	});

	it.each(others)('%s defines every English key and nothing else', (locale) => {
		const dictionary = TEAMS_MESSAGES[locale];
		expect(missingMessageKeys(english, dictionary)).toEqual([]);
		expect(missingMessageKeys(dictionary, english)).toEqual([]);
	});

	it.each(others)('%s keeps the placeholders of every English message', (locale) => {
		const dictionary = TEAMS_MESSAGES[locale];
		const mismatched = Object.keys(english).filter(
			(key) =>
				placeholders(dictionary[key] ?? '').join() !== placeholders(english[key] ?? '').join(),
		);
		expect(mismatched).toEqual([]);
	});

	it.each(TEAMS_LOCALES)('%s has no empty text, no em-dash and no stray markup', (locale) => {
		for (const [key, text] of Object.entries(TEAMS_MESSAGES[locale])) {
			expect(text.trim(), key).not.toBe('');
			expect(text, key).not.toContain(String.fromCharCode(0x2014));
			expect(text, key).not.toMatch(/[<>]/u);
		}
	});

	it('gives every plural message a one and an other form', () => {
		const names = new Set(
			Object.keys(english)
				.filter((key) => key.endsWith('.other'))
				.map((key) => key.slice(0, -'.other'.length)),
		);
		expect(names.size).toBeGreaterThan(5);
		for (const name of names) expect(english[`${name}.one`], name).toBeDefined();
	});

	it('actually translates: no other locale repeats the English wording wholesale', () => {
		for (const locale of others) {
			const same = Object.keys(english).filter(
				(key) => TEAMS_MESSAGES[locale][key] === english[key],
			);
			// Brand names, units and placeholders-only messages may match; sentences may not.
			// "Details (optional)" is also the German wording.
			const sentences = same.filter(
				(key) =>
					/\s/u.test(english[key] ?? '') &&
					!/[\d{]/u.test(english[key] ?? '') &&
					!(locale === 'de' && key === 'meeting.detailsOptional'),
			);
			expect(sentences, locale).toEqual([]);
		}
	});
});

describe('teamsI18n', () => {
	it('maps region and script tags to the shipped locales, else English', () => {
		expect(normalizeTeamsLocale('de-AT')).toBe('de');
		expect(normalizeTeamsLocale('zh-Hans')).toBe('zh-CN');
		expect(normalizeTeamsLocale('fr_CA')).toBe('fr');
		expect(normalizeTeamsLocale('zh-TW')).toBe('en');
		expect(normalizeTeamsLocale('xx')).toBe('en');
		expect(normalizeTeamsLocale('')).toBe('en');
	});

	it('fills placeholders and falls back to English, then to the key', () => {
		expect(teamsI18n('de').t('people.remove', { name: 'Ada' })).toBe('Ada entfernen');
		expect(teamsI18n('en').t('people.remove', { name: 'Ada' })).toBe('Remove Ada');
		expect(teamsI18n().t('common.save')).toBe('Save');
		expect(teamsI18n('de').t('no.such.key' as never)).toBe('no.such.key');
	});

	it('picks plural forms by the locale rules', () => {
		expect(teamsI18n('en').tn('chat.replies', 1)).toBe('1 reply');
		expect(teamsI18n('en').tn('chat.replies', 2)).toBe('2 replies');
		expect(teamsI18n('de').tn('chat.replies', 1)).toBe('1 Antwort');
		expect(teamsI18n('fr').tn('chat.replies', 0)).toBe('0 réponse');
		expect(teamsI18n('zh-CN').tn('chat.replies', 5)).toBe('5 条回复');
	});

	it('uses the browser date conventions until a locale is asked for', () => {
		expect(teamsI18n().intl).toBeUndefined();
		expect(teamsI18n('de').intl).toBe('de');
	});
});
