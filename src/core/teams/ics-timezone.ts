// The VTIMEZONE block of an iCalendar file, derived from the platform's IANA data (`Intl`) so no
// time-zone database is bundled. The offset changes of the year before the first event are read
// off the clock; when every later year (ten are checked) repeats the same pair of changes (the
// "second Sunday of March" kind) the block carries two yearly rules, which any calendar app
// extends forever. A zone whose rules differ from year to year (or that has one change only) gets
// each change written out for those years instead. DOM-free. New code.
import { MINUTE_MS, offsetMinutes } from './time-zone';

const DAY = 86_400_000;
const CHECKED_YEARS = 10;
const DAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'] as const;

export interface ZoneTransition {
	/** The UTC instant the clocks change. */
	at: number;
	/** Offsets from UTC in minutes, before and after. */
	from: number;
	to: number;
}

const cache = new Map<string, ZoneTransition[]>();

/** The clock changes of one calendar year (UTC bounds), in order. */
export function zoneTransitions(zone: string, year: number): ZoneTransition[] {
	const key = `${zone}|${year}`;
	const hit = cache.get(key);
	if (hit) return hit;
	const out: ZoneTransition[] = [];
	const begin = Date.UTC(year, 0, 1);
	const end = Date.UTC(year + 1, 0, 1);
	let before = offsetMinutes(begin, zone);
	for (let day = begin; day < end; day += DAY) {
		const after = offsetMinutes(day + DAY, zone);
		if (after === before) continue;
		let lo = day;
		let hi = day + DAY;
		while (hi - lo > MINUTE_MS) {
			const mid = Math.floor((lo + hi) / 2 / MINUTE_MS) * MINUTE_MS;
			if (offsetMinutes(mid, zone) === before) lo = mid;
			else hi = mid;
		}
		if (hi < end) out.push({ at: hi, from: before, to: after });
		before = after;
	}
	cache.set(key, out);
	return out;
}

const pad = (n: number, width = 2): string => String(n).padStart(width, '0');

/** `+0100` / `-0430` */
export function icsOffset(minutes: number): string {
	const abs = Math.abs(minutes);
	return `${minutes < 0 ? '-' : '+'}${pad(Math.floor(abs / 60))}${pad(abs % 60)}`;
}

/** The wall time on the clock just before the change, as `20260308T020000` plus its fields. */
function onset(t: ZoneTransition): { text: string; month: number; day: number; weekday: number } {
	const d = new Date(t.at + t.from * MINUTE_MS);
	return {
		text: `${pad(d.getUTCFullYear(), 4)}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}`,
		month: d.getUTCMonth() + 1,
		day: d.getUTCDate(),
		weekday: d.getUTCDay(),
	};
}

/** `2SU`, or `-1SU` for the last Sunday of the month. */
function nthWeekday(t: ZoneTransition): string {
	const o = onset(t);
	const d = new Date(t.at + t.from * MINUTE_MS);
	const days = new Date(Date.UTC(d.getUTCFullYear(), o.month, 0)).getUTCDate();
	return `${o.day + 7 > days ? -1 : Math.ceil(o.day / 7)}${DAYS[o.weekday]}`;
}

/** What makes two years' changes the same rule. */
const ruleKey = (t: ZoneTransition): string =>
	`${onset(t).month}|${nthWeekday(t)}|${onset(t).text.slice(9)}|${t.from}|${t.to}`;

function name(zone: string, at: number): string {
	try {
		const parts = new Intl.DateTimeFormat('en-US', { timeZone: zone, timeZoneName: 'short' })
			.formatToParts(new Date(at))
			.find((p) => p.type === 'timeZoneName');
		return parts?.value.replace(/[^\w+:-]/gu, '') ?? '';
	} catch {
		return '';
	}
}

const observance = (
	zone: string,
	t: ZoneTransition,
	start: string,
	rule: string | null,
): string[] => [
	`BEGIN:${t.to > t.from ? 'DAYLIGHT' : 'STANDARD'}`,
	`DTSTART:${start}`,
	`TZOFFSETFROM:${icsOffset(t.from)}`,
	`TZOFFSETTO:${icsOffset(t.to)}`,
	...(name(zone, t.at) ? [`TZNAME:${name(zone, t.at)}`] : []),
	...(rule ? [`RRULE:${rule}`] : []),
	`END:${t.to > t.from ? 'DAYLIGHT' : 'STANDARD'}`,
];

/**
 * The content lines of a VTIMEZONE for an IANA zone, valid from the year before `firstYear`.
 * The result does not depend on the viewer's zone or the current time.
 */
export function vtimezoneLines(zone: string, firstYear: number): string[] {
	const anchor = firstYear - 1;
	const first = zoneTransitions(zone, anchor);
	const later = Array.from({ length: CHECKED_YEARS }, (_, i) =>
		zoneTransitions(zone, anchor + 1 + i),
	);
	const head = ['BEGIN:VTIMEZONE', `TZID:${zone}`];
	const tail = ['END:VTIMEZONE'];
	const start = Date.UTC(anchor, 0, 1);
	const base: ZoneTransition = {
		at: start,
		from: offsetMinutes(start, zone),
		to: offsetMinutes(start, zone),
	};
	if (first.length === 0 && later.every((y) => y.length === 0))
		return [...head, ...observance(zone, base, '19700101T000000', null), ...tail];
	const repeats =
		first.length === 2 &&
		later.every(
			(y) =>
				y.length === 2 && y.every((t, i) => ruleKey(t) === ruleKey(first[i] as ZoneTransition)),
		);
	if (repeats) {
		const lines = first.flatMap((t) => {
			const o = onset(t);
			return observance(zone, t, o.text, `FREQ=YEARLY;BYMONTH=${o.month};BYDAY=${nthWeekday(t)}`);
		});
		return [...head, ...lines, ...tail];
	}
	// Irregular rules: the offset at the start, then every change that is known.
	const all = [first, ...later].flat();
	return [
		...head,
		...observance(zone, base, '19700101T000000', null),
		...all.flatMap((t) => observance(zone, t, onset(t).text, null)),
		...tail,
	];
}
