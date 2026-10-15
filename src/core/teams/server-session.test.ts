import { describe, expect, it, vi } from 'vitest';
import { endpointUrl, parseServerConfig } from './server-config';
import { requestTeamsSession } from './server-session';

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
	it('is kept when valid and sent as ?session=', () => {
		const parsed = parseServerConfig({
			mode: 'server',
			syncUrl: 'wss://h/sync',
			sessionToken: 'v1.a.b',
		});
		expect(parsed.config.sessionToken).toBe('v1.a.b');
		expect(endpointUrl('wss://h/sync', 'room', undefined, { session: 'v1.a.b' })).toBe(
			'wss://h/sync/room?session=v1.a.b',
		);
		expect(
			parseServerConfig({ sessionToken: 'x'.repeat(2000) }).config.sessionToken,
		).toBeUndefined();
	});
});
