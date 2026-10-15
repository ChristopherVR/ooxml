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
