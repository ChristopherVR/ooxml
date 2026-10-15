import { createActivityReadState } from './activity-state';
import {
	DEFAULT_NOTIFICATION_SETTINGS,
	createNotificationSettings,
	shouldNotify,
} from './notifications';

const memory = () => {
	const values = new Map<string, string>();
	return {
		values,
		storage: {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => void values.set(key, value),
		},
	};
};

describe('activity read state', () => {
	it('starts at now, so older history is read', () => {
		const { storage } = memory();
		const state = createActivityReadState(storage, 'k', () => 1000);
		expect(state.isRead('a', 999)).toBe(true);
		expect(state.isRead('b', 1001)).toBe(false);
	});

	it('marks single items read and unread, and everything with markAll', () => {
		const { storage } = memory();
		const state = createActivityReadState(storage, 'k', () => 1000);
		expect(state.mark('b', 2000, true)).toBe(true);
		expect(state.mark('b', 2000, true)).toBe(false);
		expect(state.isRead('b', 2000)).toBe(true);
		expect(state.mark('old', 500, false)).toBe(true);
		expect(state.isRead('old', 500)).toBe(false);
		expect(state.markAll(3000)).toBe(true);
		expect(state.isRead('old', 500)).toBe(true);
		expect(state.isRead('c', 3000)).toBe(true);
		expect(state.isRead('d', 3001)).toBe(false);
		expect(state.markAll(3000)).toBe(false);
	});

	it('persists per key and survives corrupt or missing storage', () => {
		const { storage, values } = memory();
		const first = createActivityReadState(storage, 'k', () => 1000);
		first.mark('b', 2000, true);
		const again = createActivityReadState(storage, 'k', () => 9999);
		expect(again.isRead('b', 2000)).toBe(true);
		expect(again.readAt()).toBe(1000);
		values.set('k', '{nope');
		expect(createActivityReadState(storage, 'k', () => 5).readAt()).toBe(5);
		expect(() =>
			createActivityReadState(undefined, 'k').mark('x', Date.now() + 5, true),
		).not.toThrow();
	});
});

describe('notifications', () => {
	const item = { kind: 'reply' as const, read: false };

	it('announces by level, never read items and never during Do not disturb', () => {
		const all = { level: 'all', sound: true } as const;
		const mentions = { level: 'mentions', sound: true } as const;
		expect(shouldNotify(item, all, 'available')).toBe(true);
		expect(shouldNotify({ ...item, read: true }, all, 'available')).toBe(false);
		expect(shouldNotify(item, mentions, 'available')).toBe(false);
		expect(shouldNotify({ ...item, kind: 'mention' }, mentions, 'busy')).toBe(true);
		expect(shouldNotify({ ...item, kind: 'mention' }, all, 'dnd')).toBe(false);
		expect(
			shouldNotify({ ...item, kind: 'mention' }, { level: 'off', sound: true }, 'available'),
		).toBe(false);
	});

	it('stores preferences and ignores unknown values', () => {
		const { storage } = memory();
		const settings = createNotificationSettings(storage, 'n');
		expect(settings.get()).toEqual(DEFAULT_NOTIFICATION_SETTINGS);
		expect(settings.set({ level: 'mentions', sound: false })).toBe(true);
		expect(settings.set({ level: 'mentions' })).toBe(false);
		expect(settings.set({ level: 'nope' as never })).toBe(false);
		expect(createNotificationSettings(storage, 'n').get()).toEqual({
			level: 'mentions',
			sound: false,
		});
	});
});
