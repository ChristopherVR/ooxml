import { type MeetingRepeat, type MonthlyBy, repeatChoices } from 'ooxml-core/teams';

const WEEKDAYS = [
	'Sunday',
	'Monday',
	'Tuesday',
	'Wednesday',
	'Thursday',
	'Friday',
	'Saturday',
] as const;
const MONTHS = [
	'January',
	'February',
	'March',
	'April',
	'May',
	'June',
	'July',
	'August',
	'September',
	'October',
	'November',
	'December',
] as const;
const ORDINALS: Record<number, string> = {
	1: 'first',
	2: 'second',
	3: 'third',
	4: 'fourth',
	[-1]: 'last',
};

/** `the third Tuesday` / `the last Friday`. */
const weekdayRule = (ordinal: number, weekday: number): string =>
	`the ${ORDINALS[ordinal] ?? 'first'} ${WEEKDAYS[weekday] ?? ''}`;

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

/** The options of the Repeat select for a start date, with labels such as "Monthly on day 15". */
export function repeatOptions(start: number, zone: string): { value: string; label: string }[] {
	const { month, day, weekday, ordinals } = repeatChoices(start, zone);
	const options = [
		{ value: 'none', label: 'Does not repeat' },
		{ value: 'daily', label: 'Every day' },
		{ value: 'weekly', label: 'Every week' },
		{ value: 'monthly-date', label: `Every month on day ${day}` },
		...ordinals.map((o) => ({
			value: `monthly-${o}`,
			label: `Every month on ${weekdayRule(o, weekday)}`,
		})),
		{ value: 'yearly', label: `Every year on ${MONTHS[month - 1]} ${day}` },
	];
	return options;
}

/** The sentence a meeting's details show: "Repeats every month on the third Tuesday". */
export function repeatText(m: {
	repeat: MeetingRepeat;
	monthlyBy?: MonthlyBy | undefined;
	ordinal?: number | undefined;
	start: number;
	zone: string;
}): string {
	if (m.repeat === 'none') return '';
	if (m.repeat === 'daily') return 'Repeats every day';
	if (m.repeat === 'weekly') return 'Repeats every week';
	const { month, day, weekday } = repeatChoices(m.start, m.zone);
	if (m.repeat === 'yearly') return `Repeats every year on ${MONTHS[month - 1]} ${day}`;
	return m.monthlyBy === 'weekday' && m.ordinal !== undefined
		? `Repeats every month on ${weekdayRule(m.ordinal, weekday)}`
		: `Repeats every month on day ${day}`;
}
