// Owns the signed-in person's own presence: the chosen status and message with their expiry, idle
// detection (Away after a configurable time without activity) and the one timer that wakes the
// client when something changes by itself. The rules are the pure functions in `./presence`; this
// only adds a clock, timers and storage, all injectable. DOM-free. New code.
import {
	type PresenceActivity,
	type PresenceStatus,
	type StatusDuration,
	applyExpiry,
	defaultPresenceSetting,
	effectivePresence,
	expiryFor,
	isIdle,
	nextExpiry,
	normalizeIdleMinutes,
	parsePresenceSetting,
	type PresenceSetting,
	sanitizeStatusMessage,
} from './presence';
import {
	type MeetingWindow,
	inMeetingAt,
	nextMeetingBoundary,
	sameWindows,
} from './presence-meetings';

const MAX_TIMER_MS = 2 ** 31 - 1;

const sameSetting = (a: PresenceSetting, b: PresenceSetting): boolean =>
	a.status === b.status &&
	a.statusExpiresAt === b.statusExpiresAt &&
	a.message === b.message &&
	a.messageExpiresAt === b.messageExpiresAt;

export interface PresenceView {
	/** What the person chose. */
	status: PresenceStatus;
	/** What others see: the chosen status, or Busy in a meeting, or Away when idle. */
	effective: PresenceStatus;
	/** Why `effective` is what it is, when it is derived: `meeting` while Busy for a meeting. */
	activity?: PresenceActivity;
	/** When the chosen status resets to Available (epoch ms), or null. */
	statusExpiresAt: number | null;
	message: string;
	messageExpiresAt: number | null;
	idle: boolean;
	/** Minutes without activity before Away; 0 never. */
	idleMinutes: number;
}

export interface PresenceStorage {
	getItem: (key: string) => string | null;
	setItem: (key: string, value: string) => void;
}

export interface PresenceControllerOptions {
	storage?: PresenceStorage | undefined;
	storageKey: string;
	idleMinutes?: number | undefined;
	now?: () => number;
	/** Called when what others see changes (also once at start). */
	publish: (status: PresenceStatus, message: string, activity?: PresenceActivity) => void;
	/** Called whenever the view changes, including idleness and expiry. */
	changed: () => void;
	/**
	 * Tells the controller the stored setting may have changed (another tab saved a new status).
	 * Returns the unsubscribe. The browser binding listens for `storage` events on the key; core
	 * stays DOM-free. Idleness stays per tab: only the chosen status and message are shared.
	 */
	subscribe?: (onChange: () => void) => () => void;
}

export interface PresenceController {
	view: () => PresenceView;
	setStatus: (status: PresenceStatus, duration?: StatusDuration) => void;
	setMessage: (message: string, duration?: StatusDuration) => void;
	setIdleMinutes: (minutes: number) => void;
	/**
	 * The windows in which the person is in a scheduled meeting (see `meetingWindows`). While one is
	 * in progress an Available person shows Busy (in a meeting); the controller wakes at the next
	 * start or end. Call again whenever the meetings change; the same windows are a no-op.
	 */
	setMeetingWindows: (windows: readonly MeetingWindow[]) => void;
	/** Report user input (the UI throttles it); it ends idleness and defers the next idle check. */
	noteActivity: () => void;
	destroy: () => void;
}

export function createPresenceController(options: PresenceControllerOptions): PresenceController {
	const now = options.now ?? Date.now;
	const { storage, storageKey } = options;
	/** The stored setting and idle choice; the defaults when nothing usable is stored. */
	const readStored = (): { setting: PresenceSetting; idleMinutes: number } => {
		let read = {
			setting: defaultPresenceSetting(),
			idleMinutes: normalizeIdleMinutes(options.idleMinutes),
		};
		try {
			const stored = JSON.parse(storage?.getItem(storageKey) ?? 'null') as Record<
				string,
				unknown
			> | null;
			read = { ...read, setting: parsePresenceSetting(stored) };
			if (stored && options.idleMinutes === undefined && 'idleMinutes' in stored)
				read.idleMinutes = normalizeIdleMinutes(stored.idleMinutes);
		} catch {
			// A corrupt value is ignored.
		}
		return read;
	};
	let { setting, idleMinutes } = readStored();
	let lastActivity = now();
	let idle = false;
	let timer: ReturnType<typeof setTimeout> | undefined;
	let destroyed = false;
	let published = '';
	let meetingWindows: readonly MeetingWindow[] = [];
	let inMeeting = false;

	const view = (): PresenceView => {
		const t = now();
		const current = applyExpiry(setting, t);
		const shown = effectivePresence(setting, idle, t, inMeetingAt(meetingWindows, t));
		return {
			status: current.status,
			effective: shown.status,
			...(shown.activity ? { activity: shown.activity } : {}),
			statusExpiresAt: current.statusExpiresAt,
			message: current.message,
			messageExpiresAt: current.messageExpiresAt,
			idle,
			idleMinutes,
		};
	};
	const save = (): void => {
		try {
			storage?.setItem(storageKey, JSON.stringify({ ...setting, idleMinutes }));
		} catch {
			// Storage is optional.
		}
	};
	const schedule = (): void => {
		clearTimeout(timer);
		if (destroyed) return;
		const times: number[] = [];
		const expiry = nextExpiry(setting);
		if (expiry !== null) times.push(expiry);
		if (!idle && idleMinutes > 0) times.push(lastActivity + idleMinutes * 60_000);
		const boundary = nextMeetingBoundary(meetingWindows, now());
		if (boundary !== null) times.push(boundary);
		if (!times.length) return;
		const wait = Math.max(0, Math.min(...times) - now());
		timer = setTimeout(evaluate, Math.min(wait, MAX_TIMER_MS));
	};
	/** Re-derive everything from the clock; publish and notify when something changed. */
	const evaluate = (): void => {
		if (destroyed) return;
		const t = now();
		const next = applyExpiry(setting, t);
		const expired = next !== setting;
		setting = next;
		const nowIdle = isIdle(lastActivity, t, idleMinutes);
		const idleChanged = nowIdle !== idle;
		idle = nowIdle;
		const nowInMeeting = inMeetingAt(meetingWindows, t);
		const meetingChanged = nowInMeeting !== inMeeting;
		inMeeting = nowInMeeting;
		if (expired) save();
		const current = view();
		const key = `${current.effective}|${current.message}|${current.activity ?? ''}`;
		if (key !== published) {
			published = key;
			options.publish(current.effective, current.message, current.activity);
		}
		if (expired || idleChanged || meetingChanged) options.changed();
		schedule();
	};
	evaluate();

	/** Another tab saved: adopt its status, message and idle choice, keep this tab's idleness. */
	const reload = (): void => {
		if (destroyed) return;
		const next = readStored();
		const same = next.idleMinutes === idleMinutes && sameSetting(next.setting, setting);
		if (same) return;
		setting = next.setting;
		idleMinutes = next.idleMinutes;
		evaluate();
		options.changed();
	};
	const unsubscribe = options.subscribe?.(reload);

	return {
		view,
		setStatus(status, duration = 'never') {
			const t = now();
			setting = {
				...applyExpiry(setting, t),
				status,
				statusExpiresAt: status === 'available' ? null : expiryFor(duration, t),
			};
			// Choosing a status is activity: Available must not flip straight back to Away.
			lastActivity = t;
			idle = false;
			save();
			evaluate();
			options.changed();
		},
		setMessage(message, duration = 'never') {
			const t = now();
			const clean = sanitizeStatusMessage(message);
			setting = {
				...applyExpiry(setting, t),
				message: clean,
				messageExpiresAt: clean ? expiryFor(duration, t) : null,
			};
			save();
			evaluate();
			options.changed();
		},
		setIdleMinutes(minutes) {
			idleMinutes = normalizeIdleMinutes(minutes);
			lastActivity = now();
			save();
			evaluate();
			options.changed();
		},
		setMeetingWindows(windows) {
			if (destroyed || sameWindows(windows, meetingWindows)) return;
			meetingWindows = windows.map((w) => ({ start: w.start, end: w.end }));
			evaluate();
		},
		noteActivity() {
			lastActivity = now();
			if (idle) evaluate();
		},
		destroy() {
			destroyed = true;
			clearTimeout(timer);
			unsubscribe?.();
		},
	};
}
