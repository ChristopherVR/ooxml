// Pure presence rules: the Teams status set, how long a status or message lasts, and when an
// idle person shows as Away. No timers and no clock here (callers pass `now`), so every binding and
// every test gets the same answers. Wire compatibility: peers still receive the three legacy
// `availability` values (`legacyAvailability`) next to the richer `status`. DOM-free. New code.
import { stripHtmlTags } from '../collab/validation';

/** What a person can be: the three original values plus Teams' other states. */
export type PresenceStatus = 'available' | 'busy' | 'dnd' | 'brb' | 'away' | 'offline';
/** The values older clients understand. */
export type LegacyAvailability = 'available' | 'busy' | 'away';

/**
 * What a derived status is for, published next to the status. Only `meeting`: Busy because a
 * scheduled meeting is in progress. Older clients ignore the field and just see Busy.
 */
export type PresenceActivity = 'meeting';

/** The activity a peer's payload means: only the known value, and only beside Busy. */
export const activityFromWire = (raw: {
	status?: unknown;
	availability?: unknown;
	activity?: unknown;
}): PresenceActivity | undefined =>
	raw.activity === 'meeting' && statusFromWire(raw) === 'busy' ? 'meeting' : undefined;

export const PRESENCE_STATUSES: readonly PresenceStatus[] = [
	'available',
	'busy',
	'dnd',
	'brb',
	'away',
	'offline',
];

/** Labels like the Teams status menu; `offline` is "Appear offline". */
export const PRESENCE_LABELS: Record<PresenceStatus, string> = {
	available: 'Available',
	busy: 'Busy',
	dnd: 'Do not disturb',
	brb: 'Be right back',
	away: 'Appear away',
	offline: 'Appear offline',
};

/** Reset-after choices of the status menu. `never` keeps the status until changed. */
export type StatusDuration = 'never' | '30m' | '1h' | '2h' | 'today' | 'week';
export const STATUS_DURATIONS: readonly StatusDuration[] = [
	'never',
	'30m',
	'1h',
	'2h',
	'today',
	'week',
];
export const STATUS_DURATION_LABELS: Record<StatusDuration, string> = {
	never: 'Never',
	'30m': '30 minutes',
	'1h': '1 hour',
	'2h': '2 hours',
	today: 'Today',
	week: 'This week',
};

export const MAX_STATUS_MESSAGE_CHARS = 120;
export const DEFAULT_IDLE_MINUTES = 5;
export const IDLE_MINUTE_CHOICES: readonly number[] = [0, 5, 10, 15, 30];

export const isPresenceStatus = (value: unknown): value is PresenceStatus =>
	PRESENCE_STATUSES.includes(value as PresenceStatus);

export const isStatusDuration = (value: unknown): value is StatusDuration =>
	STATUS_DURATIONS.includes(value as StatusDuration);

/** The closest value an older client understands: Do not disturb is Busy, the rest away. */
export function legacyAvailability(status: PresenceStatus): LegacyAvailability {
	if (status === 'available') return 'available';
	if (status === 'busy' || status === 'dnd') return 'busy';
	return 'away';
}

/**
 * The status a peer's presence payload means. A `status` this client knows wins; otherwise the
 * legacy `availability` is used, and anything else is Available (as before).
 */
export function statusFromWire(raw: { status?: unknown; availability?: unknown }): PresenceStatus {
	if (isPresenceStatus(raw.status)) return raw.status;
	return raw.availability === 'busy' || raw.availability === 'away'
		? raw.availability
		: 'available';
}

/** One client's (tab's or device's) published presence: status, message and any activity. */
export interface ClientPresence {
	status: PresenceStatus;
	message: string;
	activity?: PresenceActivity;
}

/**
 * How present a status is when one person has several clients: the lowest rank wins. A status the
 * person set that asks not to be bothered (Do not disturb, Busy) beats Available, so a busy tab is
 * not hidden by an open one; Available beats the states that mean "not at the screen" (Be right
 * back, Away), and Appear offline loses to any client that is showing. Since the chosen status is
 * shared by every tab, the clients only differ by idleness and meetings, and this keeps the more
 * conservative answer for the person.
 */
const PRESENCE_RANK: Record<PresenceStatus, number> = {
	dnd: 0,
	busy: 1,
	available: 2,
	brb: 3,
	away: 4,
	offline: 5,
};

/**
 * Negative when `a` is the better answer. After the status, ties go to the client with a message,
 * then the alphabetically first message, then the one without an activity (a Busy the person set
 * over a Busy derived from a meeting), so the answer never depends on which tab connected first.
 */
function comparePresence(a: ClientPresence, b: ClientPresence): number {
	return (
		PRESENCE_RANK[a.status] - PRESENCE_RANK[b.status] ||
		Number(b.message !== '') - Number(a.message !== '') ||
		(a.message < b.message ? -1 : a.message > b.message ? 1 : 0) ||
		Number(a.activity !== undefined) - Number(b.activity !== undefined)
	);
}

/** The client a person shows as, or null with no clients. Independent of the order of `clients`. */
export function combinePresence(clients: readonly ClientPresence[]): ClientPresence | null {
	let best: ClientPresence | null = null;
	for (const c of clients) if (!best || comparePresence(c, best) < 0) best = c;
	return best;
}

/** The presence a peer's wire payload means (status via `statusFromWire`, message sanitized). */
export function clientPresenceFromWire(raw: {
	status?: unknown;
	availability?: unknown;
	statusMessage?: unknown;
	activity?: unknown;
}): ClientPresence {
	const activity = activityFromWire(raw);
	return {
		status: statusFromWire(raw),
		message: sanitizeStatusMessage(raw.statusMessage),
		...(activity ? { activity } : {}),
	};
}

/** A custom status message: tags and control characters removed, one line, length-capped. */
export function sanitizeStatusMessage(text: unknown): string {
	if (typeof text !== 'string') return '';
	return stripHtmlTags(text)
		.replace(/[\u0000-\u001f\u007f]/gu, ' ')
		.replace(/\s+/gu, ' ')
		.trim()
		.slice(0, MAX_STATUS_MESSAGE_CHARS);
}

/**
 * When a status or message set `now` stops applying (epoch ms), or null for `never`. `today` ends
 * at the next local midnight; `week` at the end of the local Sunday (Monday 00:00).
 */
export function expiryFor(duration: StatusDuration, now: number): number | null {
	switch (duration) {
		case 'never':
			return null;
		case '30m':
			return now + 30 * 60_000;
		case '1h':
			return now + 60 * 60_000;
		case '2h':
			return now + 2 * 60 * 60_000;
		case 'today':
		case 'week': {
			const end = new Date(now);
			end.setHours(24, 0, 0, 0);
			if (duration === 'week') {
				end.setDate(end.getDate() + ((1 - end.getDay() + 7) % 7));
			}
			return end.getTime();
		}
	}
}

/** What the person chose, before idleness and expiry are applied. */
export interface PresenceSetting {
	status: PresenceStatus;
	/** When the chosen status resets to Available; null keeps it. */
	statusExpiresAt: number | null;
	message: string;
	messageExpiresAt: number | null;
}

export const defaultPresenceSetting = (): PresenceSetting => ({
	status: 'available',
	statusExpiresAt: null,
	message: '',
	messageExpiresAt: null,
});

/** Drops the parts of a setting that have expired at `now`. Returns the same object if none did. */
export function applyExpiry(setting: PresenceSetting, now: number): PresenceSetting {
	const statusExpired = setting.statusExpiresAt !== null && now >= setting.statusExpiresAt;
	const messageExpired = setting.messageExpiresAt !== null && now >= setting.messageExpiresAt;
	if (!statusExpired && !messageExpired) return setting;
	return {
		status: statusExpired ? 'available' : setting.status,
		statusExpiresAt: statusExpired ? null : setting.statusExpiresAt,
		message: messageExpired ? '' : setting.message,
		messageExpiresAt: messageExpired ? null : setting.messageExpiresAt,
	};
}

/** True once `idleMinutes` (0 turns idle detection off) have passed since the last activity. */
export function isIdle(lastActivityAt: number, now: number, idleMinutes: number): boolean {
	return idleMinutes > 0 && now - lastActivityAt >= idleMinutes * 60_000;
}

/** What others see: the status, and the activity that status stands for, if any. */
export interface EffectivePresence {
	status: PresenceStatus;
	activity?: PresenceActivity;
}

/**
 * What others see. A status the person chose (anything but Available) always wins. Available turns
 * into Busy with the `meeting` activity while a meeting they belong to is in progress, and into Away
 * when idle otherwise. Meeting beats idle: someone in a meeting is not at the keyboard, and "Away"
 * would be wrong for them.
 */
export function effectivePresence(
	setting: PresenceSetting,
	idle: boolean,
	now: number,
	inMeeting = false,
): EffectivePresence {
	const { status } = applyExpiry(setting, now);
	if (status !== 'available') return { status };
	if (inMeeting) return { status: 'busy', activity: 'meeting' };
	return { status: idle ? 'away' : 'available' };
}

/** The status others see (see `effectivePresence`); without `inMeeting` only idleness applies. */
export const effectiveStatus = (
	setting: PresenceSetting,
	idle: boolean,
	now: number,
	inMeeting = false,
): PresenceStatus => effectivePresence(setting, idle, now, inMeeting).status;

/** The earliest moment the setting changes by itself, for scheduling one timer; null if never. */
export function nextExpiry(setting: PresenceSetting): number | null {
	const times = [setting.statusExpiresAt, setting.messageExpiresAt].filter(
		(t): t is number => t !== null,
	);
	return times.length ? Math.min(...times) : null;
}

/** Reads a stored setting back; anything malformed falls back to the default. */
export function parsePresenceSetting(raw: unknown): PresenceSetting {
	const fallback = defaultPresenceSetting();
	if (!raw || typeof raw !== 'object') return fallback;
	const r = raw as Record<string, unknown>;
	const time = (v: unknown): number | null =>
		typeof v === 'number' && Number.isFinite(v) && v > 0 ? v : null;
	return {
		status: isPresenceStatus(r.status) ? r.status : fallback.status,
		statusExpiresAt: time(r.statusExpiresAt),
		message: sanitizeStatusMessage(r.message),
		messageExpiresAt: time(r.messageExpiresAt),
	};
}

/** Idle minutes from storage or a control: one of the offered choices, else the default. */
export function normalizeIdleMinutes(value: unknown): number {
	const n = Number(value);
	return IDLE_MINUTE_CHOICES.includes(n) ? n : DEFAULT_IDLE_MINUTES;
}
