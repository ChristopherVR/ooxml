import {
	DEFAULT_IDLE_MINUTES,
	MAX_STATUS_MESSAGE_CHARS,
	type PresenceSetting,
	activityFromWire,
	applyExpiry,
	clientPresenceFromWire,
	combinePresence,
	defaultPresenceSetting,
	effectivePresence,
	effectiveStatus,
	expiryFor,
	isIdle,
	legacyAvailability,
	nextExpiry,
	normalizeIdleMinutes,
	parsePresenceSetting,
	sanitizeStatusMessage,
	statusFromWire,
} from './presence';

const setting = (patch: Partial<PresenceSetting>): PresenceSetting => ({
	...defaultPresenceSetting(),
	...patch,
});

describe('presence wire compatibility', () => {
	it('maps every status to a value older clients understand', () => {
		expect(legacyAvailability('available')).toBe('available');
		expect(legacyAvailability('busy')).toBe('busy');
		expect(legacyAvailability('dnd')).toBe('busy');
		expect(legacyAvailability('brb')).toBe('away');
		expect(legacyAvailability('away')).toBe('away');
		expect(legacyAvailability('offline')).toBe('away');
	});

	it('reads the full status, else the legacy availability, else Available', () => {
		expect(statusFromWire({ status: 'dnd', availability: 'busy' })).toBe('dnd');
		expect(statusFromWire({ availability: 'busy' })).toBe('busy');
		expect(statusFromWire({ availability: 'away' })).toBe('away');
		expect(statusFromWire({ status: 'invisible', availability: 'nonsense' })).toBe('available');
		expect(statusFromWire({})).toBe('available');
	});
});

describe('status messages and durations', () => {
	it('strips markup, control characters and caps the length', () => {
		expect(sanitizeStatusMessage('  <b>Out</b>\n for\tlunch ')).toBe('Out for lunch');
		expect(sanitizeStatusMessage('x'.repeat(500))).toHaveLength(MAX_STATUS_MESSAGE_CHARS);
		expect(sanitizeStatusMessage(42)).toBe('');
	});

	it('computes expiry for each duration in local time', () => {
		// Wednesday 15 October 2025, 10:30 local.
		const now = new Date(2025, 9, 15, 10, 30).getTime();
		expect(expiryFor('never', now)).toBeNull();
		expect(expiryFor('30m', now)).toBe(now + 30 * 60_000);
		expect(expiryFor('1h', now)).toBe(now + 3_600_000);
		expect(expiryFor('2h', now)).toBe(now + 7_200_000);
		expect(expiryFor('today', now)).toBe(new Date(2025, 9, 16, 0, 0).getTime());
		// The week ends at Sunday midnight, that is Monday 20 October 00:00.
		expect(expiryFor('week', now)).toBe(new Date(2025, 9, 20, 0, 0).getTime());
		// On a Sunday the week ends tonight; on a Monday it is a full week away.
		const sunday = new Date(2025, 9, 19, 9, 0).getTime();
		expect(expiryFor('week', sunday)).toBe(new Date(2025, 9, 20, 0, 0).getTime());
		const monday = new Date(2025, 9, 20, 9, 0).getTime();
		expect(expiryFor('week', monday)).toBe(new Date(2025, 9, 27, 0, 0).getTime());
	});

	it('expires the status and the message independently', () => {
		const s = setting({
			status: 'dnd',
			statusExpiresAt: 100,
			message: 'Heads down',
			messageExpiresAt: 200,
		});
		expect(applyExpiry(s, 99)).toBe(s);
		expect(applyExpiry(s, 100)).toMatchObject({ status: 'available', message: 'Heads down' });
		expect(applyExpiry(s, 200)).toMatchObject({
			status: 'available',
			message: '',
			messageExpiresAt: null,
		});
		expect(nextExpiry(s)).toBe(100);
		expect(nextExpiry(setting({}))).toBeNull();
	});
});

describe('idle', () => {
	it('is idle after the configured minutes, never when 0', () => {
		expect(isIdle(0, 4 * 60_000, 5)).toBe(false);
		expect(isIdle(0, 5 * 60_000, 5)).toBe(true);
		expect(isIdle(0, 10 ** 9, 0)).toBe(false);
	});

	it('turns only Available into Away', () => {
		expect(effectiveStatus(setting({}), true, 0)).toBe('away');
		expect(effectiveStatus(setting({}), false, 0)).toBe('available');
		for (const status of ['busy', 'dnd', 'brb', 'offline'] as const)
			expect(effectiveStatus(setting({ status }), true, 0)).toBe(status);
	});

	it('resets an expired status before applying idleness', () => {
		const s = setting({ status: 'busy', statusExpiresAt: 50 });
		expect(effectiveStatus(s, true, 49)).toBe('busy');
		expect(effectiveStatus(s, true, 50)).toBe('away');
		expect(effectiveStatus(s, false, 50)).toBe('available');
	});

	it('normalizes the idle setting to the offered choices', () => {
		expect(normalizeIdleMinutes(10)).toBe(10);
		expect(normalizeIdleMinutes('0')).toBe(0);
		expect(normalizeIdleMinutes(7)).toBe(DEFAULT_IDLE_MINUTES);
		expect(normalizeIdleMinutes(undefined)).toBe(DEFAULT_IDLE_MINUTES);
	});
});

describe('stored settings', () => {
	it('falls back to the default for malformed input', () => {
		expect(parsePresenceSetting(null)).toEqual(defaultPresenceSetting());
		expect(
			parsePresenceSetting({ status: 'hacker', statusExpiresAt: 'soon', message: '<i>hi</i>' }),
		).toEqual({ ...defaultPresenceSetting(), message: 'hi' });
		expect(parsePresenceSetting({ status: 'dnd', statusExpiresAt: 5 })).toMatchObject({
			status: 'dnd',
			statusExpiresAt: 5,
		});
	});
});

describe('a person with several clients', () => {
	const c = (status: Parameters<typeof legacyAvailability>[0], message = '') => ({
		status,
		message,
	});

	it('has no presence without clients', () => {
		expect(combinePresence([])).toBeNull();
	});

	it('ranks do not disturb, busy, available, be right back, away, offline', () => {
		const order = ['dnd', 'busy', 'available', 'brb', 'away', 'offline'] as const;
		for (let i = 0; i < order.length; i++)
			for (let j = i + 1; j < order.length; j++) {
				expect(combinePresence([c(order[i]!), c(order[j]!)])?.status).toBe(order[i]);
				expect(combinePresence([c(order[j]!), c(order[i]!)])?.status).toBe(order[i]);
			}
	});

	it('shows an active tab over an idle one, in either order', () => {
		expect(combinePresence([c('away'), c('available')])?.status).toBe('available');
		expect(combinePresence([c('available'), c('away')])?.status).toBe('available');
	});

	it('takes the message of the client whose status won', () => {
		const won = combinePresence([c('away', 'lunch'), c('available', 'in the lab')]);
		expect(won).toEqual({ status: 'available', message: 'in the lab' });
	});

	it('breaks a tie on the message, not on the order', () => {
		const a = c('busy', 'b');
		const b = c('busy', 'a');
		const none = c('busy');
		expect(combinePresence([a, b, none])).toEqual(b);
		expect(combinePresence([none, b, a])).toEqual(b);
		expect(combinePresence([none, none])).toEqual(none);
	});

	it('reads a wire payload through the same fallbacks as a single peer', () => {
		expect(clientPresenceFromWire({ availability: 'busy' })).toEqual({
			status: 'busy',
			message: '',
		});
		expect(clientPresenceFromWire({ status: 'brb', statusMessage: '<b>back</b>\n soon' })).toEqual({
			status: 'brb',
			message: 'back soon',
		});
	});
});

describe('meeting-driven presence', () => {
	it('shows an available person as Busy in a meeting, even when idle', () => {
		const s = setting({});
		expect(effectivePresence(s, false, 0, true)).toEqual({ status: 'busy', activity: 'meeting' });
		expect(effectivePresence(s, true, 0, true)).toEqual({ status: 'busy', activity: 'meeting' });
		expect(effectiveStatus(s, true, 0, true)).toBe('busy');
	});

	it('leaves a status the person chose alone, with no activity', () => {
		for (const status of ['busy', 'dnd', 'brb', 'away', 'offline'] as const)
			expect(effectivePresence(setting({ status }), false, 0, true)).toEqual({ status });
	});

	it('applies expiry first: an expired Busy is Available again and then in the meeting', () => {
		const s = setting({ status: 'dnd', statusExpiresAt: 50 });
		expect(effectivePresence(s, false, 49, true)).toEqual({ status: 'dnd' });
		expect(effectivePresence(s, false, 50, true)).toEqual({ status: 'busy', activity: 'meeting' });
	});

	it('keeps the old behaviour without a meeting', () => {
		expect(effectivePresence(setting({}), false, 0)).toEqual({ status: 'available' });
		expect(effectivePresence(setting({}), true, 0)).toEqual({ status: 'away' });
		expect(effectiveStatus(setting({}), true, 0)).toBe('away');
	});

	it('reads the activity from the wire only for the known value beside Busy', () => {
		expect(activityFromWire({ status: 'busy', activity: 'meeting' })).toBe('meeting');
		expect(activityFromWire({ availability: 'busy', activity: 'meeting' })).toBe('meeting');
		expect(activityFromWire({ status: 'busy', activity: 'party' })).toBeUndefined();
		expect(activityFromWire({ status: 'busy', activity: { a: 1 } })).toBeUndefined();
		expect(activityFromWire({ status: 'dnd', activity: 'meeting' })).toBeUndefined();
		expect(activityFromWire({ status: 'available', activity: 'meeting' })).toBeUndefined();
		expect(activityFromWire({ status: 'busy' })).toBeUndefined();
	});

	it('prefers a Busy the person set over a Busy from a meeting, whatever the order', () => {
		const set = { status: 'busy' as const, message: '' };
		const auto = { status: 'busy' as const, message: '', activity: 'meeting' as const };
		expect(combinePresence([auto, set])).toBe(set);
		expect(combinePresence([set, auto])).toBe(set);
		expect(clientPresenceFromWire({ status: 'busy', activity: 'meeting' })).toEqual(auto);
	});
});
