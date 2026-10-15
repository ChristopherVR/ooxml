// Sign in to a Teams server: exchange a credential for a signed session that binds one user id.
// The reference server speaks this at `POST <origin>/auth/session`; a host with its own identity
// provider can instead mint a session its server accepts and pass it as `config.sessionToken`.
// The credential and the session are secrets: they are sent once and never stored or logged here.
// New code.
import type { TeamsServerConfig } from './server-config';

export interface TeamsSession {
	/** Pass as `config.sessionToken`. */
	token: string;
	/** The id the server authenticated: use it as the workspace user id. */
	userId: string;
	/** Epoch milliseconds; request a new session before then. */
	expiresAt: number;
}

export interface RequestSessionOptions {
	signal?: AbortSignal;
	fetch?: typeof fetch;
}

/** Ask the server of `config` for a session. Rejects with a generic message on any refusal. */
export async function requestTeamsSession(
	config: TeamsServerConfig,
	credentials: { userId: string; credential?: string },
	options: RequestSessionOptions = {},
): Promise<TeamsSession> {
	const base = config.syncUrl ?? config.signalingUrl;
	if (config.mode !== 'server' || !base) throw new Error('Configure a server before signing in');
	const target = new URL(base);
	target.protocol = target.protocol === 'wss:' ? 'https:' : 'http:';
	const response = await (options.fetch ?? globalThis.fetch)(`${target.origin}/auth/session`, {
		method: 'POST',
		redirect: 'error',
		credentials: 'omit',
		headers: {
			'Content-Type': 'application/json',
			...(config.token ? { Authorization: `Bearer ${config.token}` } : {}),
		},
		body: JSON.stringify({
			user: credentials.userId,
			...(credentials.credential === undefined ? {} : { credential: credentials.credential }),
		}),
		...(options.signal ? { signal: options.signal } : {}),
	});
	if (!response.ok) throw new Error(`Sign-in failed (${response.status})`);
	const body = (await response.json()) as Record<string, unknown>;
	if (
		typeof body.session !== 'string' ||
		typeof body.user !== 'string' ||
		typeof body.expiresAt !== 'number'
	)
		throw new Error('Server returned an invalid session');
	return { token: body.session, userId: body.user, expiresAt: body.expiresAt };
}
