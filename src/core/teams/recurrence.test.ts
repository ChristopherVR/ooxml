import * as Y from 'yjs';
import { occurrenceAt, occurrencesBetween } from './calendar';
import { createMeetingStore, type Meeting } from './meetings';
import { seriesIndexOf } from './recurrence';
import { DAY_MS, HOUR_MS, minuteOfDay, zonedToUtc } from './time-zone';

const NY = 'America/New_York';
const at = (month: number, day: number, hour = 9, zone = NY): number =>
	zonedToUtc({ year: 2026, month, day, hour }, zone);
const meeting = (over: Partial<Meeting> = {}): Meeting => ({
	id: 'm1',
	title: 'Stand-up',
	description: '',
	organizerId: 'ada',
	organizerName: 'Ada',
	zone: NY,
	start: at(3, 2),
	end: at(3, 2, 9) + HOUR_MS / 2,
	repeat: 'daily',
	attendees: [],
	createdAt: 0,
	cancelled: false,
	...over,
});
const starts = (m: Meeting, from: number, to: number): number[] =>
	occurrencesBetween(m, from, to).map((o) => o.start);

describe('recurrence end', () => {
	it('stops a daily series after N occurrences', () => {
		const m = meeting({ count: 3 });
		expect(starts(m, 0, 1e15)).toEqual([at(3, 2), at(3, 3), at(3, 4)]);
		expect(seriesIndexOf(m, at(3, 5))).toBeNull();
		expect(seriesIndexOf(m, at(3, 4))).toBe(2);
	});

	it('stops a weekly series on a date, keeping an occurrence that starts on it', () => {
		const m = meeting({ repeat: 'weekly', until: at(3, 16, 9) });
		expect(starts(m, 0, 1e15)).toEqual([at(3, 2), at(3, 9), at(3, 16)]);
		const before = meeting({ repeat: 'weekly', until: at(3, 16, 9) - 1 });
		expect(starts(before, 0, 1e15)).toEqual([at(3, 2), at(3, 9)]);
	});

	it('counts occurrences, not days, across the spring-forward change', () => {
		const m = meeting({ count: 10 });
		const all = occurrencesBetween(m, 0, 1e15);
		expect(all).toHaveLength(10);
		for (const o of all) expect(minuteOfDay(o.start, NY)).toBe(9 * 60);
		// 8 March 2026 is 23 hours long in New York.
		expect(all[6]!.start - all[5]!.start).toBe(DAY_MS - HOUR_MS);
	});

	it('handles a count of one and an end before the first occurrence', () => {
		expect(starts(meeting({ count: 1 }), 0, 1e15)).toEqual([at(3, 2)]);
		expect(starts(meeting({ until: at(3, 1) }), 0, 1e15)).toEqual([]);
	});

	it('ignores an end on a meeting that does not repeat', () => {
		expect(starts(meeting({ repeat: 'none', count: 2 }), 0, 1e15)).toEqual([at(3, 2)]);
	});

	it('keeps the wall time across the autumn change for a weekly series', () => {
		const m = meeting({ repeat: 'weekly', start: at(10, 26), end: at(10, 26) + HOUR_MS, count: 3 });
		const all = occurrencesBetween(m, 0, 1e15);
		expect(all.map((o) => minuteOfDay(o.start, NY))).toEqual([540, 540, 540]);
		// the clocks go back on 1 November 2026, after the first occurrence.
		expect(all[1]!.start - all[0]!.start).toBe(7 * DAY_MS + HOUR_MS);
		expect(all[2]!.start - all[1]!.start).toBe(7 * DAY_MS);
	});
});

describe('occurrence exceptions', () => {
	const cancelled = (originalStart: number) => ({ originalStart, cancelled: true });

	it('skips a cancelled occurrence and keeps the rest', () => {
		const m = meeting({ count: 4, exceptions: [cancelled(at(3, 3))] });
		expect(starts(m, 0, 1e15)).toEqual([at(3, 2), at(3, 4), at(3, 5)]);
	});

	it('applies a changed time, title and details to one occurrence', () => {
		const m = meeting({
			exceptions: [
				{
					originalStart: at(3, 3),
					cancelled: false,
					start: at(3, 3, 14),
					end: at(3, 3, 15),
					title: 'Moved',
					description: 'New room',
				},
			],
		});
		const day = occurrencesBetween(m, at(3, 3, 0), at(3, 4, 0));
		expect(day).toHaveLength(1);
		expect(day[0]).toMatchObject({
			start: at(3, 3, 14),
			end: at(3, 3, 15),
			title: 'Moved',
			description: 'New room',
			recurrenceId: at(3, 3),
			changed: true,
			index: 1,
		});
		expect(occurrenceAt(m, at(3, 3, 14))?.key).toBe(`m1@${at(3, 3, 14)}`);
		// The original slot is gone, the neighbours are untouched.
		expect(occurrenceAt(m, at(3, 3))).toBeNull();
		expect(occurrenceAt(m, at(3, 4))).toMatchObject({ index: 2 });
		expect(occurrenceAt(m, at(3, 4))?.title).toBeUndefined();
	});

	it('shows an occurrence moved to another day only on that day', () => {
		const m = meeting({
			exceptions: [
				{ originalStart: at(3, 3), cancelled: false, start: at(3, 10, 12), end: at(3, 10, 13) },
			],
		});
		expect(starts(m, at(3, 3, 0), at(3, 4, 0))).toEqual([]);
		expect(starts(m, at(3, 10, 0), at(3, 11, 0))).toEqual([at(3, 10, 9), at(3, 10, 12)]);
	});

	it('keeps a changed occurrence that ends the series', () => {
		const m = meeting({
			count: 2,
			exceptions: [{ originalStart: at(3, 3), cancelled: false, title: 'Last' }],
		});
		expect(occurrencesBetween(m, 0, 1e15).map((o) => o.title)).toEqual([undefined, 'Last']);
	});

	it('drops an exception whose original start is not in the series', () => {
		const m = meeting({
			count: 2,
			exceptions: [{ originalStart: at(3, 20), cancelled: false, title: 'Ghost' }],
		});
		expect(starts(m, 0, 1e15)).toEqual([at(3, 2), at(3, 3)]);
	});
});

describe('meeting store exceptions', () => {
	const ada = { id: 'ada', name: 'Ada' };
	const make = () => {
		const doc = new Y.Doc();
		const store = createMeetingStore(doc, ada);
		const m = store.create({
			title: 'Sync',
			start: at(3, 2),
			end: at(3, 2) + HOUR_MS,
			repeat: 'daily',
			zone: NY,
			count: 5,
		})!;
		return { doc, store, m };
	};

	it('stores a count or an end date, never both', () => {
		const { store, m } = make();
		expect(m.count).toBe(5);
		const dated = store.update(m.id, { until: at(3, 9) })!;
		expect(dated.until).toBe(at(3, 9));
		expect(dated.count).toBeUndefined();
		const counted = store.update(m.id, { count: 3 })!;
		expect(counted.count).toBe(3);
		expect(counted.until).toBeUndefined();
		const never = store.update(m.id, { count: null })!;
		expect(never.count).toBeUndefined();
		expect(never.until).toBeUndefined();
		expect(store.update(m.id, { until: at(3, 1) })).toBeNull();
	});

	it('cancels and edits single occurrences, rejecting times that are not occurrences', () => {
		const { store, m } = make();
		expect(store.cancelOccurrence(m.id, at(3, 3))).toBe(true);
		expect(store.cancelOccurrence(m.id, at(3, 3, 10))).toBe(false);
		expect(store.cancelOccurrence(m.id, at(3, 20))).toBe(false);
		const edited = store.editOccurrence(m.id, at(3, 4), {
			title: '<b>Late</b> sync',
			start: at(3, 4, 15),
			end: at(3, 4, 16),
		})!;
		expect(edited.exceptions).toEqual([
			{ originalStart: at(3, 3), cancelled: true },
			{
				originalStart: at(3, 4),
				cancelled: false,
				start: at(3, 4, 15),
				end: at(3, 4, 16),
				title: 'Late sync',
			},
		]);
		expect(
			store.editOccurrence(m.id, at(3, 4), { start: at(3, 4, 16), end: at(3, 4, 15) }),
		).toBeNull();
		expect(store.editOccurrence(m.id, at(3, 5), { title: ' ' })).toBeNull();
		const again = store.editOccurrence(m.id, at(3, 4), { description: 'Bring notes' })!;
		expect(again.exceptions?.[1]).toMatchObject({ title: 'Late sync', description: 'Bring notes' });
	});

	it('only lets the organizer change occurrences, and syncs them to other peers', () => {
		const { doc, store, m } = make();
		store.cancelOccurrence(m.id, at(3, 3));
		const other = new Y.Doc();
		Y.applyUpdate(other, Y.encodeStateAsUpdate(doc));
		const bob = createMeetingStore(other, { id: 'bob', name: 'Bob' });
		expect(bob.get(m.id)?.exceptions).toHaveLength(1);
		expect(bob.cancelOccurrence(m.id, at(3, 4))).toBe(false);
		expect(bob.editOccurrence(m.id, at(3, 4), { title: 'x' })).toBeNull();
	});

	it('voids exceptions when the series is rescheduled or stops repeating', () => {
		const { store, m } = make();
		store.cancelOccurrence(m.id, at(3, 3));
		expect(store.update(m.id, { title: 'Renamed' })!.exceptions).toHaveLength(1);
		const moved = store.update(m.id, { start: at(3, 2, 10), end: at(3, 2, 11) })!;
		expect(moved.exceptions).toBeUndefined();
		store.cancelOccurrence(m.id, at(3, 3, 10));
		const single = store.update(m.id, { repeat: 'none' })!;
		expect(single.exceptions).toBeUndefined();
		expect(single.count).toBeUndefined();
	});

	it('ignores malformed stored exceptions', () => {
		const { doc, m } = make();
		const row = doc.getMap<Y.Map<unknown>>('teams:meetings').get(m.id)!;
		const bad = new Y.Map<unknown>();
		bad.set('nope', { cancelled: true });
		bad.set(String(at(3, 3)), { start: 5 });
		bad.set(String(at(3, 4)), 'text');
		bad.set(String(at(3, 5)), { cancelled: true });
		row.set('exceptions', bad);
		const read = createMeetingStore(doc, ada).get(m.id)!;
		expect(read.exceptions).toEqual([{ originalStart: at(3, 5), cancelled: true }]);
	});
});
