// Notification preferences and the rule that decides whether an activity item is announced. The
// preferences are personal (stored on the device, not in Yjs). Delivery itself (a toast, the
// browser's Notification API, a sound) is the UI's job; the core only decides *whether* and emits
// the item. Do not disturb suppresses everything, as in Teams. DOM-free.
import type { ActivityItem } from './activity';
import type { PresenceStatus } from './presence';
import type { StorageLike } from './store';

/** `all`: every kind of activity. `mentions`: only when you are mentioned. `off`: nothing. */
export type NotificationLevel = 'all' | 'mentions' | 'off';

export interface NotificationSettings {
	level: NotificationLevel;
	sound: boolean;
}

export const NOTIFICATION_LEVELS: readonly NotificationLevel[] = ['all', 'mentions', 'off'];
export const DEFAULT_NOTIFICATION_SETTINGS: NotificationSettings = { level: 'all', sound: true };

export const isNotificationLevel = (value: unknown): value is NotificationLevel =>
	NOTIFICATION_LEVELS.includes(value as NotificationLevel);

/** Whether `item` is announced to a person with these settings and this (effective) status. */
export function shouldNotify(
	item: Pick<ActivityItem, 'kind' | 'read'>,
	settings: NotificationSettings,
	status: PresenceStatus,
): boolean {
	if (item.read || settings.level === 'off' || status === 'dnd') return false;
	// A meeting reminder is one you asked for, so only Off and Do not disturb silence it.
	return settings.level === 'all' || item.kind === 'mention' || item.kind === 'meeting';
}

export interface NotificationSettingsStore {
	get: () => NotificationSettings;
	/** Returns whether anything changed. Unknown values are ignored. */
	set: (next: Partial<NotificationSettings>) => boolean;
}

export function createNotificationSettings(
	storage: StorageLike | undefined,
	key: string,
): NotificationSettingsStore {
	let settings: NotificationSettings = { ...DEFAULT_NOTIFICATION_SETTINGS };
	try {
		const raw = storage?.getItem(key);
		const stored: unknown = raw && raw.length <= 1024 ? JSON.parse(raw) : null;
		if (stored && typeof stored === 'object') {
			const { level, sound } = stored as Record<string, unknown>;
			settings = {
				level: isNotificationLevel(level) ? level : settings.level,
				sound: typeof sound === 'boolean' ? sound : settings.sound,
			};
		}
	} catch {
		/* Defaults remain usable without storage. */
	}
	return {
		get: () => ({ ...settings }),
		set(next) {
			const level = isNotificationLevel(next.level) ? next.level : settings.level;
			const sound = typeof next.sound === 'boolean' ? next.sound : settings.sound;
			if (level === settings.level && sound === settings.sound) return false;
			settings = { level, sound };
			try {
				storage?.setItem(key, JSON.stringify(settings));
			} catch {
				/* In-memory only. */
			}
			return true;
		},
	};
}
