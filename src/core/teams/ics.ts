// Export a meeting as an iCalendar (RFC 5545) file. Times are written in UTC, so any calendar app
// shows them in the reader's zone; a repeating meeting becomes one event with an RRULE, which
// calendars expand on their own. An exported repeat is anchored to UTC, so a calendar that does not
// know the organiser's zone shifts it by an hour across daylight saving. There are no attendee
// addresses in a workspace without accounts, so names go in the description. DOM-free. New code.
import type { Meeting } from './meetings';

const pad = (n: number): string => String(n).padStart(2, '0');

/** `20261015T140000Z` */
export function icsUtc(ms: number): string {
	const d = new Date(ms);
	return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

export const icsEscape = (text: string): string =>
	text
		.replace(/\\/gu, '\\\\')
		.replace(/\r?\n/gu, '\\n')
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

/** The `.ics` text (CRLF line endings) for one meeting. */
export function meetingToIcs(meeting: Meeting, options: IcsOptions = {}): string {
	const people = [
		`Organizer: ${meeting.organizerName}`,
		...(meeting.attendees.length
			? [`Attendees: ${meeting.attendees.map((a) => a.name).join(', ')}`]
			: []),
	];
	const description = [meeting.description, ...people].filter(Boolean).join('\n\n');
	const lines = [
		'BEGIN:VCALENDAR',
		'VERSION:2.0',
		'PRODID:-//OpenTeams//Calendar//EN',
		'CALSCALE:GREGORIAN',
		'METHOD:PUBLISH',
		'BEGIN:VEVENT',
		`UID:${meeting.id}@openteams`,
		`DTSTAMP:${icsUtc(options.now ?? Date.now())}`,
		`DTSTART:${icsUtc(meeting.start)}`,
		`DTEND:${icsUtc(meeting.end)}`,
		`SUMMARY:${icsEscape(meeting.title)}`,
		`DESCRIPTION:${icsEscape(description)}`,
		...(meeting.repeat === 'daily' ? ['RRULE:FREQ=DAILY'] : []),
		...(meeting.repeat === 'weekly' ? ['RRULE:FREQ=WEEKLY'] : []),
		...(options.url ? [`URL:${options.url}`] : []),
		`STATUS:${meeting.cancelled ? 'CANCELLED' : 'CONFIRMED'}`,
		'BEGIN:VALARM',
		'ACTION:DISPLAY',
		`DESCRIPTION:${icsEscape(meeting.title)}`,
		'TRIGGER:-PT5M',
		'END:VALARM',
		'END:VEVENT',
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
