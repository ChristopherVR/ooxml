import { createPresenceController, type PresenceController } from './presence-controller';
import type { PresenceStatus } from './presence';

const MIN = 60_000;

function make(opts: { idleMinutes?: number; stored?: string } = {}) {
	const store = new Map<string, string>();
	if (opts.stored) store.set('k', opts.stored);
	const published: string[] = [];
	let changes = 0;
	const controller: PresenceController = createPresenceController({
		storage: { getItem: (k) => store.get(k) ?? null, setItem: (k, v) => void store.set(k, v) },
		storageKey: 'k',
		idleMinutes: opts.idleMinutes,
		publish: (status: PresenceStatus, message: string) => published.push(`${status}|${message}`),
		changed: () => changes++,
	});
	return { controller, published, store, changes: () => changes };
}

beforeEach(() => {
	vi.useFakeTimers();
	vi.setSystemTime(new Date(2025, 9, 15, 10, 0));
});
afterEach(() => vi.useRealTimers());

describe('presence controller', () => {
	it('publishes Available at start and Away after the idle time', () => {
		const { controller, published, changes } = make({ idleMinutes: 5 });
		expect(published).toEqual(['available|']);
		vi.advanceTimersByTime(4 * MIN);
		expect(controller.view().idle).toBe(false);
		vi.advanceTimersByTime(MIN);
		expect(controller.view()).toMatchObject({ idle: true, effective: 'away', status: 'available' });
		expect(published).toEqual(['available|', 'away|']);
		expect(changes()).toBeGreaterThan(0);
		controller.destroy();
	});

	it('returns to Available on activity and defers the idle check', () => {
		const { controller, published } = make({ idleMinutes: 5 });
		vi.advanceTimersByTime(4 * MIN);
		controller.noteActivity();
		vi.advanceTimersByTime(4 * MIN);
		expect(controller.view().idle).toBe(false);
		vi.advanceTimersByTime(MIN);
		expect(controller.view().idle).toBe(true);
		controller.noteActivity();
		expect(controller.view()).toMatchObject({ idle: false, effective: 'available' });
		expect(published).toEqual(['available|', 'away|', 'available|']);
		controller.destroy();
	});

	it('never goes idle when the setting is 0, and does not override Do not disturb', () => {
		const off = make({ idleMinutes: 0 });
		vi.advanceTimersByTime(60 * MIN);
		expect(off.controller.view().effective).toBe('available');
		off.controller.destroy();
		const dnd = make({ idleMinutes: 5 });
		dnd.controller.setStatus('dnd');
		vi.advanceTimersByTime(30 * MIN);
		expect(dnd.controller.view()).toMatchObject({ idle: true, effective: 'dnd' });
		dnd.controller.destroy();
	});

	it('resets a timed status to Available and clears an expiring message', () => {
		const { controller, published } = make({ idleMinutes: 0 });
		controller.setStatus('brb', '1h');
		controller.setMessage('Lunch', '30m');
		expect(published.at(-1)).toBe('brb|Lunch');
		vi.advanceTimersByTime(30 * MIN);
		expect(controller.view()).toMatchObject({ message: '', effective: 'brb' });
		expect(published.at(-1)).toBe('brb|');
		vi.advanceTimersByTime(30 * MIN);
		expect(controller.view()).toMatchObject({ status: 'available', statusExpiresAt: null });
		expect(published.at(-1)).toBe('available|');
		controller.destroy();
	});

	it('keeps Available without an expiry and clears the message with an empty string', () => {
		const { controller } = make({ idleMinutes: 0 });
		controller.setStatus('available', '1h');
		expect(controller.view().statusExpiresAt).toBeNull();
		controller.setMessage('Hi', 'today');
		expect(controller.view().messageExpiresAt).not.toBeNull();
		controller.setMessage('');
		expect(controller.view()).toMatchObject({ message: '', messageExpiresAt: null });
		controller.destroy();
	});

	it('remembers the choice and the idle setting, dropping what expired', () => {
		const first = make({ idleMinutes: 10 });
		first.controller.setStatus('offline', '2h');
		first.controller.setMessage('Away', 'never');
		first.controller.destroy();
		const saved = first.store.get('k')!;
		const again = make({ stored: saved });
		expect(again.controller.view()).toMatchObject({
			status: 'offline',
			message: 'Away',
			idleMinutes: 10,
		});
		again.controller.destroy();
		vi.advanceTimersByTime(3 * 60 * MIN);
		const later = make({ stored: saved });
		expect(later.controller.view()).toMatchObject({ status: 'available', message: 'Away' });
		later.controller.destroy();
	});

	it('changes the idle setting and restarts the countdown', () => {
		const { controller } = make({ idleMinutes: 5 });
		controller.setIdleMinutes(15);
		vi.advanceTimersByTime(14 * MIN);
		expect(controller.view().idle).toBe(false);
		vi.advanceTimersByTime(MIN);
		expect(controller.view().idle).toBe(true);
		controller.setIdleMinutes(0);
		expect(controller.view()).toMatchObject({ idle: false, idleMinutes: 0 });
		controller.destroy();
	});
});

/** Two tabs of one person: a shared fake storage that tells the other tab when one saves. */
function twoTabs(idleMinutes = 5) {
	const store = new Map<string, string>();
	const listeners = new Set<{ tab: number; fn: () => void }>();
	const tab = (id: number) => {
		const published: string[] = [];
		let changes = 0;
		const controller = createPresenceController({
			storage: {
				getItem: (k) => store.get(k) ?? null,
				setItem: (k, v) => {
					store.set(k, v);
					for (const l of [...listeners]) if (l.tab !== id) l.fn();
				},
			},
			storageKey: 'k',
			idleMinutes,
			publish: (status, message) => published.push(`${status}|${message}`),
			changed: () => changes++,
			subscribe: (fn) => {
				const entry = { tab: id, fn };
				listeners.add(entry);
				return () => listeners.delete(entry);
			},
		});
		return { controller, published, changes: () => changes, listeners: () => listeners.size };
	};
	return { a: tab(1), b: tab(2), listeners: () => listeners.size };
}

describe('a status chosen in another tab', () => {
	it('reaches this tab and is published from it', () => {
		const { a, b } = twoTabs(0);
		a.controller.setStatus('dnd', '1h');
		a.controller.setMessage('Deep work');
		expect(b.controller.view()).toMatchObject({
			status: 'dnd',
			effective: 'dnd',
			message: 'Deep work',
		});
		expect(b.published.at(-1)).toBe('dnd|Deep work');
		expect(b.changes()).toBeGreaterThan(0);
		a.controller.setStatus('available');
		expect(b.controller.view().effective).toBe('available');
		expect(b.published.at(-1)).toBe('available|Deep work');
		a.controller.destroy();
		b.controller.destroy();
	});

	it('keeps idleness per tab', () => {
		const { a, b } = twoTabs(5);
		vi.advanceTimersByTime(5 * MIN);
		expect(b.controller.view().idle).toBe(true);
		// Tab A is where the person is: choosing a status there is activity in A only.
		a.controller.setStatus('busy');
		expect(a.controller.view().idle).toBe(false);
		expect(b.controller.view()).toMatchObject({ idle: true, status: 'busy', effective: 'busy' });
		a.controller.setStatus('available');
		expect(a.controller.view().effective).toBe('available');
		expect(b.controller.view()).toMatchObject({
			idle: true,
			status: 'available',
			effective: 'away',
		});
		a.controller.destroy();
		b.controller.destroy();
	});

	it('ignores a notification that changes nothing', () => {
		const { a, b } = twoTabs(0);
		a.controller.setStatus('brb');
		const before = b.changes();
		const published = b.published.length;
		a.controller.setIdleMinutes(0);
		expect(b.changes()).toBe(before);
		expect(b.published.length).toBe(published);
		a.controller.destroy();
		b.controller.destroy();
	});

	it('stops listening on destroy', () => {
		const { a, b, listeners } = twoTabs(0);
		expect(listeners()).toBe(2);
		b.controller.destroy();
		expect(listeners()).toBe(1);
		a.controller.setStatus('dnd');
		expect(b.controller.view().status).toBe('available');
		a.controller.destroy();
		expect(listeners()).toBe(0);
	});
});

describe('presence in a meeting', () => {
	const window = (fromMin: number, toMin: number) => ({
		start: Date.now() + fromMin * MIN,
		end: Date.now() + toMin * MIN,
	});

	it('is Busy (in a meeting) from the start to the end of the window and wakes by itself', () => {
		const { controller, published } = make({ idleMinutes: 0 });
		controller.setMeetingWindows([window(10, 40)]);
		expect(published).toEqual(['available|']);
		vi.advanceTimersByTime(10 * MIN);
		expect(controller.view()).toMatchObject({
			effective: 'busy',
			activity: 'meeting',
			status: 'available',
		});
		expect(published.at(-1)).toBe('busy|');
		vi.advanceTimersByTime(29 * MIN);
		expect(controller.view().effective).toBe('busy');
		vi.advanceTimersByTime(MIN);
		expect(controller.view().effective).toBe('available');
		expect(controller.view().activity).toBeUndefined();
		expect(published).toEqual(['available|', 'busy|', 'available|']);
		controller.destroy();
	});

	it('tells the wire why, and asks to be shown again when the view changes', () => {
		const calls: (string | undefined)[] = [];
		let changes = 0;
		const controller = createPresenceController({
			storageKey: 'k',
			idleMinutes: 0,
			publish: (_s, _m, activity) => calls.push(activity),
			changed: () => changes++,
		});
		controller.setMeetingWindows([window(-5, 5)]);
		expect(calls).toEqual([undefined, 'meeting']);
		expect(changes).toBeGreaterThan(0);
		vi.advanceTimersByTime(5 * MIN);
		expect(calls).toEqual([undefined, 'meeting', undefined]);
		controller.destroy();
	});

	it('lets a status the person chose win, and returns to Busy-in-meeting when it resets', () => {
		const { controller } = make({ idleMinutes: 0 });
		controller.setMeetingWindows([window(-5, 120)]);
		expect(controller.view().effective).toBe('busy');
		controller.setStatus('brb', '30m');
		expect(controller.view()).toMatchObject({ effective: 'brb' });
		expect(controller.view().activity).toBeUndefined();
		vi.advanceTimersByTime(30 * MIN);
		expect(controller.view()).toMatchObject({ effective: 'busy', activity: 'meeting' });
		controller.destroy();
	});

	it('stays Busy when idle during the meeting, and goes Away if idle after it', () => {
		const { controller } = make({ idleMinutes: 5 });
		controller.setMeetingWindows([window(2, 20)]);
		vi.advanceTimersByTime(10 * MIN);
		expect(controller.view()).toMatchObject({ idle: true, effective: 'busy', activity: 'meeting' });
		vi.advanceTimersByTime(10 * MIN);
		expect(controller.view()).toMatchObject({ idle: true, effective: 'away' });
		controller.destroy();
	});

	it('ignores an unchanged list, and a list set after destroy', () => {
		const { controller, published, changes } = make({ idleMinutes: 0 });
		const w = [window(-1, 10)];
		controller.setMeetingWindows(w);
		const n = published.length;
		const c = changes();
		controller.setMeetingWindows(w.map((x) => ({ ...x })));
		expect(published.length).toBe(n);
		expect(changes()).toBe(c);
		controller.destroy();
		controller.setMeetingWindows([]);
		expect(controller.view().effective).toBe('busy');
	});

	it('drops the Busy at once when the windows are cleared (meeting cancelled)', () => {
		const { controller, published } = make({ idleMinutes: 0 });
		controller.setMeetingWindows([window(-1, 30)]);
		expect(controller.view().effective).toBe('busy');
		controller.setMeetingWindows([]);
		expect(controller.view().effective).toBe('available');
		expect(published.at(-1)).toBe('available|');
		controller.destroy();
	});
});
