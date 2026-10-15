// Time-zone arithmetic on instants (UTC milliseconds) with the platform's IANA data, no library:
// wall-clock parts of an instant in a zone, the instant of a wall-clock time in a zone (DST gaps
// move forward, repeated hours take the first), and day starts and day steps that keep the local
// time of day across daylight saving changes. Calendar code stores UTC and renders in a zone.
// DOM-free. New code.

export interface ZonedParts {
	year: number;
	/** 1-12. */
	month: number;
	day: number;
	hour: number;
	minute: number;
	second: number;
	/** 0 = Sunday. */
	weekday: number;
}

export const MINUTE_MS = 60_000;
export const HOUR_MS = 3_600_000;
export const DAY_MS = 86_400_000;

const formatters = new Map<string, Intl.DateTimeFormat>();
const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

export function isValidZone(zone: unknown): zone is string {
	if (typeof zone !== 'string' || zone.length === 0 || zone.length > 64) return false;
	try {
		return Boolean(formatter(zone));
	} catch {
		return false;
	}
}

/** The viewer's zone, UTC when the platform does not say. */
export function localZone(): string {
	try {
		const zone = new Intl.DateTimeFormat().resolvedOptions().timeZone;
		return isValidZone(zone) ? zone : 'UTC';
	} catch {
		return 'UTC';
	}
}

function formatter(zone: string): Intl.DateTimeFormat {
	let f = formatters.get(zone);
	if (!f) {
		f = new Intl.DateTimeFormat('en-US', {
			timeZone: zone,
			hourCycle: 'h23',
			year: 'numeric',
			month: 'numeric',
			day: 'numeric',
			hour: 'numeric',
			minute: 'numeric',
			second: 'numeric',
			weekday: 'short',
		});
		formatters.set(zone, f);
	}
	return f;
}

export function zonedParts(ms: number, zone: string): ZonedParts {
	const out: Record<string, string> = {};
	for (const part of formatter(zone).formatToParts(new Date(ms))) out[part.type] = part.value;
	return {
		year: Number(out.year),
		month: Number(out.month),
		day: Number(out.day),
		hour: Number(out.hour) % 24,
		minute: Number(out.minute),
		second: Number(out.second),
		weekday: Math.max(0, WEEKDAYS.indexOf(out.weekday ?? '')),
	};
}

/** Minutes the zone is ahead of UTC at that instant. */
export function offsetMinutes(ms: number, zone: string): number {
	const p = zonedParts(ms, zone);
	const asUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
	return Math.round((asUtc - Math.floor(ms / 1000) * 1000) / MINUTE_MS);
}

export type WallTime = Pick<ZonedParts, 'year' | 'month' | 'day'> &
	Partial<Pick<ZonedParts, 'hour' | 'minute' | 'second'>>;

/** The instant at which a zone's clocks read this wall time. Out-of-range fields roll over. */
export function zonedToUtc(wall: WallTime, zone: string): number {
	const naive = Date.UTC(
		wall.year,
		wall.month - 1,
		wall.day,
		wall.hour ?? 0,
		wall.minute ?? 0,
		wall.second ?? 0,
	);
	const first = naive - offsetMinutes(naive, zone) * MINUTE_MS;
	const second = naive - offsetMinutes(first, zone) * MINUTE_MS;
	if (first === second) return first;
	// Across a transition the two guesses differ: a gap has no answer (use the later offset), a
	// repeated hour has two (the earlier instant).
	const exact = [first, second].filter((c) => zonedParts(c, zone).hour === (wall.hour ?? 0));
	return exact.length ? Math.min(...exact) : Math.max(first, second);
}

/** The instant of local midnight of the day containing `ms`. */
export function startOfDay(ms: number, zone: string): number {
	const p = zonedParts(ms, zone);
	return zonedToUtc({ year: p.year, month: p.month, day: p.day }, zone);
}

/** The same local time of day, `days` later (or earlier) in the zone. */
export function addDays(ms: number, days: number, zone: string): number {
	const p = zonedParts(ms, zone);
	return zonedToUtc({ ...p, day: p.day + days }, zone);
}

/** The local wall time of an instant as a sortable key such as `2026-10-15`. */
export function dayKey(ms: number, zone: string): string {
	const p = zonedParts(ms, zone);
	return `${p.year}-${String(p.month).padStart(2, '0')}-${String(p.day).padStart(2, '0')}`;
}

/** Minutes after local midnight, so a day with a DST change still lines up with the clock. */
export function minuteOfDay(ms: number, zone: string): number {
	const p = zonedParts(ms, zone);
	return p.hour * 60 + p.minute;
}
