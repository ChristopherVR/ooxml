import { teamsI18n } from '../i18n/translator';
import { repeatOptions, repeatText } from './repeat-text';

/** The third Tuesday of October 2026. */
const start = Date.UTC(2026, 9, 20, 12);
const thirdTuesday = {
	repeat: 'monthly' as const,
	monthlyBy: 'weekday' as const,
	ordinal: 3,
	start,
	zone: 'UTC',
};

describe('repeat sentences by locale', () => {
	it.each([
		['en', 'Repeats every month on the third Tuesday'],
		['de', 'Wiederholt sich: jeden Monat am dritten Dienstag'],
		['es', 'Se repite: cada mes el tercer martes'],
		['fr', 'Se répète : tous les mois, le troisième mardi'],
		['zh-CN', '重复：每月第三个星期二'],
	])('%s: %s', (locale, sentence) => {
		expect(repeatText(thirdTuesday, teamsI18n(locale))).toBe(sentence);
	});

	it('words the last weekday, the month day and the yearly date in each language', () => {
		const last = { ...thirdTuesday, ordinal: -1 };
		expect(repeatText(last, teamsI18n('de'))).toBe(
			'Wiederholt sich: jeden Monat am letzten Dienstag',
		);
		expect(repeatText({ repeat: 'monthly', start, zone: 'UTC' }, teamsI18n('fr'))).toBe(
			'Se répète : tous les mois, le 20',
		);
		expect(repeatText({ repeat: 'yearly', start, zone: 'UTC' }, teamsI18n('de'))).toBe(
			'Wiederholt sich: jedes Jahr am 20. Oktober',
		);
		expect(repeatText({ repeat: 'yearly', start, zone: 'UTC' }, teamsI18n('es'))).toBe(
			'Se repite: cada año el 20 de octubre',
		);
		expect(repeatText({ repeat: 'daily', start, zone: 'UTC' }, teamsI18n('zh-CN'))).toBe(
			'重复：每天',
		);
	});

	it('lists the options of the Repeat select in the display language', () => {
		const labels = repeatOptions(start, 'UTC', teamsI18n('de')).map((o) => o.label);
		expect(labels).toEqual([
			'Wiederholt sich nicht',
			'Jeden Tag',
			'Jede Woche',
			'Jeden Monat am 20.',
			'Jeden Monat am dritten Dienstag',
			'Jedes Jahr am 20. Oktober',
		]);
	});
});
