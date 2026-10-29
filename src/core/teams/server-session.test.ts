import { describe, expect, it, vi } from 'vitest';
import { connectionAuth, endpointUrl, parseServerConfig } from './server-config';
import {
	TeamsSignInError,
	checkTeamsSession,
	probeTeamsAccounts,
	requestTeamsSession,
	signOutTeamsSession,
} from './server-session';

const config = parseServerConfig({
	mode: 'server',
	syncUrl: 'wss://chat.example.com/sync',
	token: 'shared',
}).config;

describe('requestTeamsSession', () => {
	it('posts the credential to the server origin and returns the session', async () => {
		const fetchMock = vi.fn(
			async () =>
				new Response(JSON.stringify({ session: 's.t.u', user: 'alice@corp', expiresAt: 5 })),
		);
		const session = await requestTeamsSession(
			config,
			{ userId: 'alice', credential: 'pw' },
			{ fetch: fetchMock as unknown as typeof fetch },
		);
		expect(session).toEqual({ token: 's.t.u', userId: 'alice@corp', expiresAt: 5 });
		const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
		expect(url).toBe('https://chat.example.com/auth/session');
		expect(init.credentials).toBe('omit');
		expect((init.headers as Record<string, string>).Authorization).toBe('Bearer shared');
		expect(JSON.parse(init.body as string)).toEqual({ user: 'alice', credential: 'pw' });
	});

	it('rejects a refusal without echoing the credential', async () => {
		const fetchMock = vi.fn(async () => new Response('no', { status: 401 }));
		const failure = requestTeamsSession(
			config,
			{ userId: 'alice', credential: 'hunter2' },
			{ fetch: fetchMock as unknown as typeof fetch },
		);
		await expect(failure).rejects.toThrow('Sign-in failed (401)');
		await expect(failure).rejects.not.toThrow(/hunter2/u);
	});

	it('rejects a malformed answer and a local config', async () => {
		const bad = vi.fn(async () => new Response(JSON.stringify({ session: 1 })));
		await expect(
			requestTeamsSession(config, { userId: 'a' }, { fetch: bad as unknown as typeof fetch }),
		).rejects.toThrow('invalid session');
		await expect(
			requestTeamsSession({ mode: 'local', iceServers: [] }, { userId: 'a' }),
		).rejects.toThrow('Configure a server');
	});
});

describe('session token in the config', () => {
	it('is kept when valid and travels as a subprotocol', () => {
		const parsed = parseServerConfig({
			mode: 'server',
			syncUrl: 'wss://h/sync',
			sessionToken: 'v1.a.b',
		});
		expect(parsed.config.sessionToken).toBe('v1.a.b');
		expect(connectionAuth(parsed.config)).toEqual({
			params: {},
			protocols: ['openteams', 'openteams.session.v1.a.b'],
		});
		expect(endpointUrl('wss://h/sync', 'room', undefined, { session: 'v1.a.b' })).toBe(
			'wss://h/sync/room?session=v1.a.b',
		);
		expect(
			parseServerConfig({ sessionToken: 'x'.repeat(2000) }).config.sessionToken,
		).toBeUndefined();
	});
});

describe('sign-in failures', () => {
	const failing = (status: number, headers: Record<string, string> = {}) =>
		vi.fn(async () => new Response('x', { status, headers })) as unknown as typeof fetch;
	const reason = async (fetchMock: typeof fetch) =>
		requestTeamsSession(config, { userId: 'a', credential: 'b' }, { fetch: fetchMock }).catch(
			(error: unknown) => error as TeamsSignInError,
		);

	it('tells a wrong credential, a rate limit and a server error apart', async () => {
		expect(await reason(failing(401))).toMatchObject({ reason: 'credentials' });
		expect(await reason(failing(500))).toMatchObject({ reason: 'server' });
		const limited = await reason(failing(429, { 'Retry-After': '12' }));
		expect(limited).toMatchObject({ reason: 'rate-limited', retryAfterSec: 12 });
		expect(limited).toBeInstanceOf(TeamsSignInError);
	});

	it('reports an unreachable server without the underlying message', async () => {
		const down = vi.fn(async () => {
			throw new TypeError('connect ECONNREFUSED 10.0.0.1');
		}) as unknown as typeof fetch;
		const error = await reason(down);
		expect(error).toMatchObject({ reason: 'network', message: 'Could not reach the server' });
	});
});

describe('probe, check and sign out', () => {
	const answer = (status: number, body = '') =>
		vi.fn(async () => new Response(body || null, { status })) as unknown as typeof fetch;

	it('reads whether the server requires accounts', async () => {
		expect(
			await probeTeamsAccounts(config, { fetch: answer(200, '{"ok":true,"accounts":true}') }),
		).toBe(true);
		expect(await probeTeamsAccounts(config, { fetch: answer(200, '{"ok":true}') })).toBe(false);
		expect(await probeTeamsAccounts(config, { fetch: answer(500) })).toBeNull();
		expect(await probeTeamsAccounts({ mode: 'local', iceServers: [] })).toBeNull();
		const down = vi.fn(async () => {
			throw new Error('offline');
		}) as unknown as typeof fetch;
		expect(await probeTeamsAccounts(config, { fetch: down })).toBeNull();
	});

	it('maps the answer to valid, rejected or unreachable', async () => {
		expect(await checkTeamsSession(config, 's', { fetch: answer(204) })).toBe('valid');
		expect(await checkTeamsSession(config, 's', { fetch: answer(401) })).toBe('rejected');
		expect(await checkTeamsSession(config, 's', { fetch: answer(503) })).toBe('unreachable');
	});

	it('sends the session in the body with the workspace token, never in the URL', async () => {
		const fetchMock = answer(204);
		await signOutTeamsSession(config, 'sess.ion', { fetch: fetchMock });
		const [url, init] = (fetchMock as unknown as { mock: { calls: [string, RequestInit][] } }).mock
			.calls[0]!;
		expect(url).toBe('https://chat.example.com/auth/signout');
		expect(url).not.toContain('sess.ion');
		expect(JSON.parse(init.body as string)).toEqual({ session: 'sess.ion' });
		expect((init.headers as Record<string, string>).Authorization).toBe('Bearer shared');
	});
});
