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

/** Why a sign-in did not complete; `retryAfterSec` accompanies `rate-limited`. */
export type SignInFailure = 'credentials' | 'rate-limited' | 'network' | 'server';

export class TeamsSignInError extends Error {
	readonly reason: SignInFailure;
	readonly retryAfterSec: number | undefined;
	constructor(reason: SignInFailure, message: string, retryAfterSec?: number) {
		super(message);
		this.name = 'TeamsSignInError';
		this.reason = reason;
		this.retryAfterSec = retryAfterSec;
	}
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
	const origin = serverOrigin(config);
	if (!origin) throw new Error('Configure a server before signing in');
	let response: Response;
	try {
		response = await sessionFetch(options)(`${origin}/auth/session`, {
			method: 'POST',
			redirect: 'error',
			credentials: 'omit',
			headers: jsonHeaders(config),
			body: JSON.stringify({
				user: credentials.userId,
				...(credentials.credential === undefined ? {} : { credential: credentials.credential }),
			}),
			...(options.signal ? { signal: options.signal } : {}),
		});
	} catch (error) {
		if (options.signal?.aborted) throw error;
		throw new TeamsSignInError('network', 'Could not reach the server');
	}
	if (!response.ok) {
		const wait = Number(response.headers.get('Retry-After'));
		throw new TeamsSignInError(
			response.status === 429
				? 'rate-limited'
				: response.status === 401 || response.status === 400
					? 'credentials'
					: 'server',
			`Sign-in failed (${response.status})`,
			response.status === 429 && Number.isFinite(wait) && wait > 0 ? Math.ceil(wait) : undefined,
		);
	}
	const body = (await response.json()) as Record<string, unknown>;
	if (
		typeof body.session !== 'string' ||
		typeof body.user !== 'string' ||
		typeof body.expiresAt !== 'number'
	)
		throw new Error('Server returned an invalid session');
	return { token: body.session, userId: body.user, expiresAt: body.expiresAt };
}

/** `https://host` for the HTTP side of a server config (the origin of its sync or signaling URL). */
export function serverOrigin(config: TeamsServerConfig): string | null {
	const base = config.syncUrl ?? config.signalingUrl;
	if (config.mode !== 'server' || !base) return null;
	const target = new URL(base);
	target.protocol = target.protocol === 'wss:' ? 'https:' : 'http:';
	return target.origin;
}

const sessionFetch = (options: RequestSessionOptions): typeof fetch =>
	options.fetch ?? globalThis.fetch.bind(globalThis);

const jsonHeaders = (config: TeamsServerConfig): Record<string, string> => ({
	'Content-Type': 'application/json',
	...(config.token ? { Authorization: `Bearer ${config.token}` } : {}),
});

/**
 * Whether the server requires accounts (`/health` reports `accounts`). Null when it cannot be
 * asked (unreachable, not a Teams server): callers then carry on as before, without a sign-in.
 */
export async function probeTeamsAccounts(
	config: TeamsServerConfig,
	options: RequestSessionOptions = {},
): Promise<boolean | null> {
	const origin = serverOrigin(config);
	if (!origin) return null;
	try {
		const response = await sessionFetch(options)(`${origin}/health`, {
			redirect: 'error',
			credentials: 'omit',
			...(options.signal ? { signal: options.signal } : {}),
		});
		if (!response.ok) return null;
		const body = (await response.json()) as Record<string, unknown>;
		return typeof body.accounts === 'boolean' ? body.accounts : false;
	} catch {
		return null;
	}
}

/** `valid`, `rejected` (expired, revoked or forged: sign in again) or `unreachable` (try later). */
export type SessionCheck = 'valid' | 'rejected' | 'unreachable';

export async function checkTeamsSession(
	config: TeamsServerConfig,
	session: string,
	options: RequestSessionOptions = {},
): Promise<SessionCheck> {
	return postSession(config, '/auth/check', session, options);
}

/** End every session of the signed-in user on the server. Best effort: resolves on any outcome. */
export async function signOutTeamsSession(
	config: TeamsServerConfig,
	session: string,
	options: RequestSessionOptions = {},
): Promise<SessionCheck> {
	return postSession(config, '/auth/signout', session, options);
}

async function postSession(
	config: TeamsServerConfig,
	path: string,
	session: string,
	options: RequestSessionOptions,
): Promise<SessionCheck> {
	const origin = serverOrigin(config);
	if (!origin) return 'unreachable';
	try {
		const response = await sessionFetch(options)(`${origin}${path}`, {
			method: 'POST',
			redirect: 'error',
			credentials: 'omit',
			headers: jsonHeaders(config),
			body: JSON.stringify({ session }),
			...(options.signal ? { signal: options.signal } : {}),
		});
		if (response.ok) return 'valid';
		return response.status === 401 ? 'rejected' : 'unreachable';
	} catch {
		return 'unreachable';
	}
}
