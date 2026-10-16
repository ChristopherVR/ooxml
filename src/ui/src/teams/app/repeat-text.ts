import { type MeetingRepeat, type MonthlyBy, repeatChoices } from 'ooxml-core/teams';
import { teamsI18n, type TeamsI18n, type TeamsKey } from '../i18n/translator';

/** Month and day as the locale writes them ("October 15"); the year is a leap year so 29 Feb exists. */
const monthDay = (month: number, day: number, i18n: TeamsI18n): string =>
	new Intl.DateTimeFormat(i18n.locale, { month: 'long', day: 'numeric', timeZone: 'UTC' }).format(
		Date.UTC(2000, month - 1, day),
	);

/** A weekday's name in the display locale (0 is Sunday); names sit inside translated sentences. */
const weekdayName = (weekday: number, i18n: TeamsI18n): string =>
	new Intl.DateTimeFormat(i18n.locale, { weekday: 'long', timeZone: 'UTC' }).format(
		Date.UTC(2000, 0, 2 + (weekday % 7)),
	);

const ORDINAL_KEYS: Record<number, TeamsKey> = {
	1: 'repeat.monthlyOrdinal.1',
	2: 'repeat.monthlyOrdinal.2',
	3: 'repeat.monthlyOrdinal.3',
	4: 'repeat.monthlyOrdinal.4',
	[-1]: 'repeat.monthlyOrdinal.last',
};

/** `Every month on the third Tuesday` / `Every month on the last Friday`, in the locale. */
const monthlyWeekday = (ordinal: number, weekday: number, i18n: TeamsI18n): string =>
	i18n.t(ORDINAL_KEYS[ordinal] ?? 'repeat.monthlyOrdinal.1', {
		weekday: weekdayName(weekday, i18n),
	});

/** The repeat the dialog's select shows: one option per way a start date can repeat. */
export interface RepeatSelection {
	repeat: MeetingRepeat;
	monthlyBy: MonthlyBy;
	/** With `monthlyBy: 'weekday'`. */
	ordinal: number;
}

export const NO_REPEAT: RepeatSelection = { repeat: 'none', monthlyBy: 'date', ordinal: 1 };

/** The select value of a repeat: `none`, `daily`, `weekly`, `monthly-date`, `monthly-3`, `yearly`. */
export const repeatValue = (r: RepeatSelection): string =>
	r.repeat === 'monthly'
		? r.monthlyBy === 'weekday'
			? `monthly-${r.ordinal}`
			: 'monthly-date'
		: r.repeat;

export function parseRepeatValue(value: string): RepeatSelection {
	if (value === 'monthly-date') return { repeat: 'monthly', monthlyBy: 'date', ordinal: 1 };
	const m = /^monthly-(-?\d)$/u.exec(value);
	if (m) return { repeat: 'monthly', monthlyBy: 'weekday', ordinal: Number(m[1]) };
	if (value === 'daily' || value === 'weekly' || value === 'yearly')
		return { repeat: value, monthlyBy: 'date', ordinal: 1 };
	return NO_REPEAT;
}

/** The options of the Repeat select for a start date, with labels such as "Every month on day 15". */
export function repeatOptions(
	start: number,
	zone: string,
	i18n: TeamsI18n = teamsI18n(),
): { value: string; label: string }[] {
	const { month, day, weekday, ordinals } = repeatChoices(start, zone);
	return [
		{ value: 'none', label: i18n.t('repeat.none') },
		{ value: 'daily', label: i18n.t('repeat.daily') },
		{ value: 'weekly', label: i18n.t('repeat.weekly') },
		{ value: 'monthly-date', label: i18n.t('repeat.monthlyDay', { day }) },
		...ordinals.map((o) => ({
			value: `monthly-${o}`,
			label: monthlyWeekday(o, weekday, i18n),
		})),
		{ value: 'yearly', label: i18n.t('repeat.yearlyOn', { date: monthDay(month, day, i18n) }) },
	];
}

/** Makes an option label read inside a sentence ("Every day" -> "every day"). */
const inSentence = (label: string, i18n: TeamsI18n): string =>
	label.charAt(0).toLocaleLowerCase(i18n.locale) + label.slice(1);

/** The sentence a meeting's details show: "Repeats every month on the third Tuesday". */
export function repeatText(
	m: {
		repeat: MeetingRepeat;
		monthlyBy?: MonthlyBy | undefined;
		ordinal?: number | undefined;
		start: number;
		zone: string;
	},
	i18n: TeamsI18n = teamsI18n(),
): string {
	if (m.repeat === 'none') return '';
	const { month, day, weekday } = repeatChoices(m.start, m.zone);
	const rule =
		m.repeat === 'daily'
			? i18n.t('repeat.daily')
			: m.repeat === 'weekly'
				? i18n.t('repeat.weekly')
				: m.repeat === 'yearly'
					? i18n.t('repeat.yearlyOn', { date: monthDay(month, day, i18n) })
					: m.monthlyBy === 'weekday' && m.ordinal !== undefined
						? monthlyWeekday(m.ordinal, weekday, i18n)
						: i18n.t('repeat.monthlyDay', { day });
	return i18n.t('repeat.sentence', { rule: inSentence(rule, i18n) });
}
