// Export a meeting as an iCalendar (RFC 5545) file. The start and end carry the organiser's zone
// (TZID) with a VTIMEZONE block derived from the platform's zone data, so a repeating meeting keeps
// its wall-clock time across daylight saving in any calendar app; a UTC organiser uses plain UTC
// times. A repeat is one event with an RRULE (COUNT or UNTIL), cancelled occurrences are EXDATEs and
// a changed occurrence is an overriding event with the same UID and a RECURRENCE-ID. There are no
// attendee addresses in a workspace without accounts, so names go in the description. DOM-free.
// New code.
import { vtimezoneLines } from './ics-timezone';
import type { Meeting } from './meetings';
import { zonedParts } from './time-zone';

const pad = (n: number, width = 2): string => String(n).padStart(width, '0');

/** `20261015T140000Z` */
export function icsUtc(ms: number): string {
	const d = new Date(ms);
	return `${pad(d.getUTCFullYear(), 4)}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

export const icsEscape = (text: string): string =>
	text
		.replace(/\\/gu, '\\\\')
		.replace(/\r\n|\r|\n/gu, '\\n')
		.replace(/;/gu, '\\;')
		.replace(/,/gu, '\\,');

/** Fold a content line at 75 octets, continuing with a space (RFC 5545 section 3.1). */
export function icsFold(line: string): string[] {
	const encoder = new TextEncoder();
	const out: string[] = [];
	let current = '';
	let bytes = 0;
	for (const ch of line) {
		const size = encoder.encode(ch).length;
		const limit = out.length === 0 ? 75 : 74;
		if (bytes + size > limit) {
			out.push(current);
			current = '';
			bytes = 0;
		}
		current += ch;
		bytes += size;
	}
	out.push(current);
	return out.map((part, i) => (i === 0 ? part : ` ${part}`));
}

export interface IcsOptions {
	now?: number;
	/** Shown in the calendar app's event link; omitted when empty. */
	url?: string;
}

const UTC_ZONES = new Set(['UTC', 'Etc/UTC', 'GMT', 'Etc/GMT']);

/** `20261015T090000`: the wall time of an instant in a zone, without a suffix. */
export function icsLocal(ms: number, zone: string): string {
	const p = zonedParts(ms, zone);
	return `${pad(p.year, 4)}${pad(p.month)}${pad(p.day)}T${pad(p.hour)}${pad(p.minute)}${pad(p.second)}`;
}

/** A date-time property: `DTSTART;TZID=America/New_York:20261015T090000`, or `...:...Z` in UTC. */
export function icsDateTime(name: string, ms: number, zone: string): string {
	return UTC_ZONES.has(zone)
		? `${name}:${icsUtc(ms)}`
		: `${name};TZID=${zone}:${icsLocal(ms, zone)}`;
}

function alarm(title: string): string[] {
	return [
		'BEGIN:VALARM',
		'ACTION:DISPLAY',
		`DESCRIPTION:${icsEscape(title)}`,
		'TRIGGER:-PT5M',
		'END:VALARM',
	];
}

function recurrenceRule(meeting: Meeting): string[] {
	if (meeting.repeat === 'none') return [];
	const freq = meeting.repeat === 'daily' ? 'DAILY' : 'WEEKLY';
	// UNTIL is in UTC when DTSTART names a zone (RFC 5545 section 3.3.10).
	const end =
		meeting.count !== undefined
			? `;COUNT=${meeting.count}`
			: meeting.until !== undefined
				? `;UNTIL=${icsUtc(meeting.until)}`
				: '';
	const utc = UTC_ZONES.has(meeting.zone);
	const exdates = (meeting.exceptions ?? [])
		.filter((e) => e.cancelled)
		.map((e) => (utc ? icsUtc(e.originalStart) : icsLocal(e.originalStart, meeting.zone)));
	return [
		`RRULE:FREQ=${freq}${end}`,
		...(exdates.length
			? [utc ? `EXDATE:${exdates.join(',')}` : `EXDATE;TZID=${meeting.zone}:${exdates.join(',')}`]
			: []),
	];
}

/** The `.ics` text (CRLF line endings) for one meeting, with its changed occurrences. */
export function meetingToIcs(meeting: Meeting, options: IcsOptions = {}): string {
	const people = [
		`Organizer: ${meeting.organizerName}`,
		...(meeting.attendees.length
			? [`Attendees: ${meeting.attendees.map((a) => a.name).join(', ')}`]
			: []),
	];
	const details = (text: string): string =>
		`DESCRIPTION:${icsEscape([text, ...people].filter(Boolean).join('\n\n'))}`;
	const uid = `UID:${meeting.id}@openteams`;
	const stamp = `DTSTAMP:${icsUtc(options.now ?? Date.now())}`;
	const status = `STATUS:${meeting.cancelled ? 'CANCELLED' : 'CONFIRMED'}`;
	const link = options.url ? [`URL:${options.url}`] : [];
	const changed =
		meeting.repeat === 'none' ? [] : (meeting.exceptions ?? []).filter((e) => !e.cancelled);
	const years = [
		meeting.start,
		...changed.flatMap((e) => [e.originalStart, e.start ?? e.originalStart]),
	].map((t) => zonedParts(t, meeting.zone).year);
	const lines = [
		'BEGIN:VCALENDAR',
		'VERSION:2.0',
		'PRODID:-//OpenTeams//Calendar//EN',
		'CALSCALE:GREGORIAN',
		'METHOD:PUBLISH',
		...(UTC_ZONES.has(meeting.zone) ? [] : vtimezoneLines(meeting.zone, Math.min(...years))),
		'BEGIN:VEVENT',
		uid,
		stamp,
		icsDateTime('DTSTART', meeting.start, meeting.zone),
		icsDateTime('DTEND', meeting.end, meeting.zone),
		`SUMMARY:${icsEscape(meeting.title)}`,
		details(meeting.description),
		...recurrenceRule(meeting),
		...link,
		status,
		...alarm(meeting.title),
		'END:VEVENT',
		...changed.flatMap((e) => {
			const start = e.start ?? e.originalStart;
			const end = e.end ?? start + (meeting.end - meeting.start);
			const title = e.title ?? meeting.title;
			return [
				'BEGIN:VEVENT',
				uid,
				stamp,
				icsDateTime('RECURRENCE-ID', e.originalStart, meeting.zone),
				icsDateTime('DTSTART', start, meeting.zone),
				icsDateTime('DTEND', end, meeting.zone),
				`SUMMARY:${icsEscape(title)}`,
				details(e.description ?? meeting.description),
				...link,
				status,
				...alarm(title),
				'END:VEVENT',
			];
		}),
		'END:VCALENDAR',
	];
	return `${lines.flatMap(icsFold).join('\r\n')}\r\n`;
}

/** A safe download name such as `Weekly-sync.ics`. */
export const icsFileName = (meeting: Pick<Meeting, 'title'>): string =>
	`${
		meeting.title
			.replace(/[^\w.-]+/gu, '-')
			.replace(/^-+|-+$/gu, '')
			.slice(0, 60) || 'meeting'
	}.ics`;
