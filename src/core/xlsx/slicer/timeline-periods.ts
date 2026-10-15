export const TIMELINE_LEVELS = ['years', 'quarters', 'months', 'days'] as const;

/** One selectable period of a timeline (a year, quarter, month or day). */
export interface TimelinePeriod {
	index: number;
	/** The text of its button (`2025`, `Q1`, `Feb`, `3`). */
	label: string;
	/** The parent period shown above it (`2025`, `Feb 2025`). */
	group: string;
	/** First and last day, as `YYYY-MM-DD`; the last day is inclusive. */
	first: string;
	last: string;
}

export interface TimelineSelection {
	start: string;
	end: string;
}

const MAX_PERIODS = 20_000;
const DAY = 86_400_000;
const pad = (value: number, width = 2) => String(value).padStart(width, '0');
const ymd = (time: number): string => {
	const d = new Date(time);
	return `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
};
const parse = (iso: string): number => {
	const [y, m, d] = iso.slice(0, 10).split('-').map(Number);
	return Date.UTC(y ?? 1970, (m ?? 1) - 1, d ?? 1);
};

/** `YYYY-MM-DDT00:00:00`, the form timeline caches store. */
export const toIsoDate = (day: string): string => `${day.slice(0, 10)}T00:00:00`;

function monthName(time: number, locale: string): string {
	try {
		return new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' }).format(time);
	} catch {
		return new Intl.DateTimeFormat('en', { month: 'short', timeZone: 'UTC' }).format(time);
	}
}

/**
 * Every period of `level` (0 years, 1 quarters, 2 months, 3 days) between the bounds; the end
 * bound is exclusive, as Excel stores it (Jan 1 of the year after the data).
 */
export function timelinePeriods(
	level: number,
	boundsStart: string,
	boundsEnd: string,
	locale = 'en',
): TimelinePeriod[] {
	const start = parse(boundsStart);
	const end = parse(boundsEnd);
	const out: TimelinePeriod[] = [];
	const first = new Date(start);
	for (
		let cursor = Date.UTC(first.getUTCFullYear(), level === 3 ? first.getUTCMonth() : 0, 1);
		cursor < end && out.length < MAX_PERIODS;
	) {
		const d = new Date(cursor);
		const year = d.getUTCFullYear();
		const month = d.getUTCMonth();
		let next: number;
		let label: string;
		let group: string;
		if (level === 0) {
			next = Date.UTC(year + 1, 0, 1);
			label = group = String(year);
		} else if (level === 1) {
			next = Date.UTC(year, month + 3, 1);
			label = `Q${Math.floor(month / 3) + 1}`;
			group = String(year);
		} else if (level === 2) {
			next = Date.UTC(year, month + 1, 1);
			label = monthName(cursor, locale);
			group = String(year);
		} else {
			next = cursor;
			// Days are generated month by month below.
			label = '';
			group = '';
		}
		if (level === 3) {
			const monthEnd = Date.UTC(year, month + 1, 1);
			for (let day = cursor; day < monthEnd && out.length < MAX_PERIODS; day += DAY) {
				if (day >= start && day < end)
					out.push({
						index: out.length,
						label: String(new Date(day).getUTCDate()),
						group: `${monthName(cursor, locale)} ${year}`,
						first: ymd(day),
						last: ymd(day),
					});
			}
			cursor = monthEnd;
			continue;
		}
		if (next > start)
			out.push({ index: out.length, label, group, first: ymd(cursor), last: ymd(next - DAY) });
		cursor = next;
	}
	return out;
}

/** The selection covering periods `a` to `b` (either order). */
export function selectionOfPeriods(
	periods: readonly TimelinePeriod[],
	a: number,
	b: number,
): TimelineSelection | undefined {
	const from = periods[Math.min(a, b)];
	const to = periods[Math.max(a, b)];
	return from && to ? { start: toIsoDate(from.first), end: toIsoDate(to.last) } : undefined;
}

/** The first and last period a selection covers, or undefined when it covers none. */
export function periodsInSelection(
	periods: readonly TimelinePeriod[],
	selection: TimelineSelection | undefined,
): [first: number, last: number] | undefined {
	if (!selection) return undefined;
	const start = selection.start.slice(0, 10);
	const end = selection.end.slice(0, 10);
	const inside = periods.filter((period) => period.last >= start && period.first <= end);
	const a = inside[0];
	const b = inside.at(-1);
	return a && b ? [a.index, b.index] : undefined;
}

/** `Feb 2025 - Apr 2025`: the selection as the header shows it. */
export function selectionLabel(
	periods: readonly TimelinePeriod[],
	selection: TimelineSelection | undefined,
	level: number,
	locale = 'en',
): string {
	const range = periodsInSelection(periods, selection);
	if (!range) return '';
	const name = (period: TimelinePeriod): string =>
		level === 0
			? period.label
			: level === 3
				? `${period.group.split(' ')[0]} ${period.label}, ${period.group.split(' ')[1]}`
				: `${period.label} ${period.group}`;
	void locale;
	const from = name(periods[range[0]]!);
	const to = name(periods[range[1]]!);
	return from === to ? from : `${from} - ${to}`;
}

/** The ISO date-time bounds Excel stores: Jan 1 of the first year to Jan 1 after the last. */
export function boundsOfSerialRange(
	minDay: string,
	maxDay: string,
): { start: string; end: string } {
	const first = Number(minDay.slice(0, 4));
	const last = Number(maxDay.slice(0, 4));
	return { start: `${pad(first, 4)}-01-01T00:00:00`, end: `${pad(last + 1, 4)}-01-01T00:00:00` };
}
