// Exceptions to a repeating meeting: one occurrence cancelled, or changed (a new time, title or
// details). Each is keyed by the occurrence's original start, which is what iCalendar calls
// RECURRENCE-ID. Stored in the meeting row as a nested map so edits to different occurrences merge.
// Values are re-validated when read back. DOM-free. New code.
import * as Y from 'yjs';
import { sanitizeChannelName, sanitizeMessageText } from './model';

export interface MeetingException {
	/** UTC start the occurrence would have had without the exception. */
	originalStart: number;
	cancelled: boolean;
	/** Overrides for a changed occurrence. */
	start?: number;
	end?: number;
	title?: string;
	description?: string;
}

export interface OccurrencePatch {
	start?: number;
	end?: number;
	title?: string;
	description?: string;
}

export const MAX_EXCEPTIONS = 400;
const MAX_SPAN = 24 * 3_600_000;
const time = (v: unknown): number | null =>
	typeof v === 'number' && Number.isFinite(v) && v > 0 && v < 8.64e15 ? Math.floor(v) : null;

/** One stored value as an exception, or null when it is not a valid one. */
export function readException(key: string, raw: unknown): MeetingException | null {
	const originalStart = time(Number(key));
	if (originalStart === null || !raw || typeof raw !== 'object') return null;
	const r = raw as Record<string, unknown>;
	if (r.cancelled === true) return { originalStart, cancelled: true };
	const start = time(r.start);
	const end = time(r.end);
	const title = typeof r.title === 'string' ? sanitizeChannelName(r.title).slice(0, 120) : '';
	const description =
		typeof r.description === 'string' ? sanitizeMessageText(r.description).slice(0, 2000) : null;
	if ((start === null) !== (end === null)) return null;
	if (start !== null && end !== null && (end <= start || end - start > MAX_SPAN)) return null;
	return {
		originalStart,
		cancelled: false,
		...(start !== null && end !== null ? { start, end } : {}),
		...(title ? { title } : {}),
		...(description !== null ? { description } : {}),
	};
}

export function readExceptions(raw: unknown): MeetingException[] {
	if (!(raw instanceof Y.Map)) return [];
	const out: MeetingException[] = [];
	raw.forEach((value, key) => {
		const e = readException(key, value);
		if (e && out.length < MAX_EXCEPTIONS) out.push(e);
	});
	return out.sort((a, b) => a.originalStart - b.originalStart);
}

/** The exception for an occurrence that started at `originalStart`, if any. */
export const exceptionAt = (
	exceptions: readonly MeetingException[] | undefined,
	originalStart: number,
): MeetingException | undefined => exceptions?.find((e) => e.originalStart === originalStart);

/** The row's exception map, created on first use. */
export function exceptionMap(row: Y.Map<unknown>): Y.Map<unknown> {
	const existing = row.get('exceptions');
	if (existing instanceof Y.Map) return existing;
	const created = new Y.Map<unknown>();
	row.set('exceptions', created);
	return created;
}

/** The plain stored value for a change (only the fields that are set). */
export function exceptionValue(
	patch: OccurrencePatch,
	previous: MeetingException | undefined,
): Record<string, unknown> {
	const merged = { ...(previous && !previous.cancelled ? previous : {}), ...patch };
	const out: Record<string, unknown> = {};
	if (merged.start !== undefined && merged.end !== undefined) {
		out.start = merged.start;
		out.end = merged.end;
	}
	if (merged.title !== undefined) out.title = sanitizeChannelName(merged.title);
	if (merged.description !== undefined) out.description = sanitizeMessageText(merged.description);
	return out;
}
