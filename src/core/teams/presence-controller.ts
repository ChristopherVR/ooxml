// Owns the signed-in person's own presence: the chosen status and message with their expiry, idle
// detection (Away after a configurable time without activity) and the one timer that wakes the
// client when something changes by itself. The rules are the pure functions in `./presence`; this
// only adds a clock, timers and storage, all injectable. DOM-free. New code.
import {
	type PresenceStatus,
	type StatusDuration,
	applyExpiry,
	defaultPresenceSetting,
	effectiveStatus,
	expiryFor,
	isIdle,
	nextExpiry,
	normalizeIdleMinutes,
	parsePresenceSetting,
	sanitizeStatusMessage,
} from './presence';

const MAX_TIMER_MS = 2 ** 31 - 1;

export interface PresenceView {
	/** What the person chose. */
	status: PresenceStatus;
	/** What others see: the chosen status, or Away when idle. */
	effective: PresenceStatus;
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
	publish: (status: PresenceStatus, message: string) => void;
	/** Called whenever the view changes, including idleness and expiry. */
	changed: () => void;
}

export interface PresenceController {
	view: () => PresenceView;
	setStatus: (status: PresenceStatus, duration?: StatusDuration) => void;
	setMessage: (message: string, duration?: StatusDuration) => void;
	setIdleMinutes: (minutes: number) => void;
	/** Report user input (the UI throttles it); it ends idleness and defers the next idle check. */
	noteActivity: () => void;
	destroy: () => void;
}

export function createPresenceController(options: PresenceControllerOptions): PresenceController {
	const now = options.now ?? Date.now;
	const { storage, storageKey } = options;
	let setting = defaultPresenceSetting();
	let idleMinutes = normalizeIdleMinutes(options.idleMinutes);
	try {
		const stored = JSON.parse(storage?.getItem(storageKey) ?? 'null') as Record<
			string,
			unknown
		> | null;
		setting = parsePresenceSetting(stored);
		if (stored && options.idleMinutes === undefined && 'idleMinutes' in stored)
			idleMinutes = normalizeIdleMinutes(stored.idleMinutes);
	} catch {
		// A corrupt value is ignored.
	}
	let lastActivity = now();
	let idle = false;
	let timer: ReturnType<typeof setTimeout> | undefined;
	let destroyed = false;
	let published = '';

	const effective = (): PresenceStatus => effectiveStatus(setting, idle, now());
	const view = (): PresenceView => {
		const current = applyExpiry(setting, now());
		return {
			status: current.status,
			effective: effective(),
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
		if (expired) save();
		const current = view();
		const key = `${current.effective}|${current.message}`;
		if (key !== published) {
			published = key;
			options.publish(current.effective, current.message);
		}
		if (expired || idleChanged) options.changed();
		schedule();
	};
	evaluate();

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
		noteActivity() {
			lastActivity = now();
			if (idle) evaluate();
		},
		destroy() {
			destroyed = true;
			clearTimeout(timer);
		},
	};
}
