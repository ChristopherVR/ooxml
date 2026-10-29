// Session tokens for the reference server. A client proves who it is once (POST /auth/session,
// through the `authenticate` hook or the built-in `users` table) and gets a signed, expiring token
// that binds one user id. Sockets then present that token, and the server uses the id inside it,
// never the `user` query parameter. Tokens are HMAC-SHA-256 over the payload; the secret never
// leaves the process and is never logged. Pure helpers, no I/O.
import { createHash, createHmac, randomBytes, timingSafeEqual } from 'node:crypto';
import { hashPassword, isHashedCredential, verifyPassword } from './passwords.mjs';

const USER = /^[\w.@:+-]{1,64}$/u;
export const DEFAULT_SESSION_TTL_MS = 12 * 60 * 60 * 1000;
const b64 = (buffer) => Buffer.from(buffer).toString('base64url');

export const ROLES = ['moderator'];
export const isRole = (value) => typeof value === 'string' && ROLES.includes(value);

export const isUserId = (value) => typeof value === 'string' && USER.test(value);

/** Constant-time string comparison that does not leak the length. */
export function sameSecret(a, b) {
	const ha = createHash('sha256').update(String(a)).digest();
	const hb = createHash('sha256').update(String(b)).digest();
	return timingSafeEqual(ha, hb);
}

/**
 * @param {{ secret?: string, ttlMs?: number, now?: () => number }} [options]
 * Without a `secret` a random one is generated, so sessions end when the process restarts.
 * Revocation: every user has a session generation (default 0) that tokens carry; `revoke(user)`
 * bumps it, which invalidates every session issued before. `generations` seeds it (persisted by
 * the caller through `onRevoke`), so revocations survive a restart that keeps the secret.
 */
export function createSessionSigner({
	secret,
	ttlMs = DEFAULT_SESSION_TTL_MS,
	now = Date.now,
	generations = {},
	onRevoke = () => {},
} = {}) {
	const generation = new Map(Object.entries(generations));
	const currentGeneration = (user) => generation.get(user) ?? 0;
	const key = createHash('sha256')
		.update(`teams-session|${secret || randomBytes(32).toString('hex')}`)
		.digest();
	const mac = (payload) => createHmac('sha256', key).update(payload).digest();
	return {
		ttlMs,
		/** Invalidate every session issued so far for `user`. Returns the new generation. */
		revoke(user) {
			if (!isUserId(user)) throw new Error('invalid user id');
			const next = currentGeneration(user) + 1;
			generation.set(user, next);
			onRevoke(user, next, Object.fromEntries(generation));
			return next;
		},
		/** A token for `user` (optionally with a `role`), valid for `ttlMs`. */
		issue(user, role) {
			if (!isUserId(user)) throw new Error('invalid user id');
			const exp = now() + ttlMs;
			const claims = {
				u: user,
				exp,
				g: currentGeneration(user),
				...(isRole(role) ? { r: role } : {}),
			};
			const payload = b64(JSON.stringify(claims));
			return { token: `v1.${payload}.${b64(mac(payload))}`, user, expiresAt: exp };
		},
		/** The authenticated user id of a token, or null when it is forged, malformed or expired. */
		verify(token) {
			return this.inspect(token)?.user ?? null;
		},
		/** `{ user, role }` of a valid token (role is undefined for an ordinary member), else null. */
		inspect(token) {
			if (typeof token !== 'string' || token.length > 1024) return null;
			const parts = token.split('.');
			if (parts.length !== 3 || parts[0] !== 'v1') return null;
			const [, payload, signature] = parts;
			const expected = mac(payload);
			const given = Buffer.from(signature, 'base64url');
			if (given.length !== expected.length || !timingSafeEqual(given, expected)) return null;
			try {
				const claims = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
				if (!isUserId(claims.u) || typeof claims.exp !== 'number' || claims.exp <= now())
					return null;
				if ((typeof claims.g === 'number' ? claims.g : 0) !== currentGeneration(claims.u))
					return null;
				return { user: claims.u, role: isRole(claims.r) ? claims.r : undefined };
			} catch {
				return null;
			}
		},
	};
}

/** A users entry is a credential string or `{ credential, role? }` (role: 'moderator'). */
function accountOf(entry) {
	if (typeof entry === 'string') return { credential: entry };
	if (entry && typeof entry === 'object')
		return { credential: entry.credential, ...(isRole(entry.role) ? { role: entry.role } : {}) };
	return { credential: undefined };
}

/**
 * The credential table behind `users`: user id -> credential. Returns an `authenticate` function.
 * A credential is either plain text or a hash made by `hashPassword` (`scrypt$...`); both are
 * compared in constant time, and an unknown user costs one hash check so response time does not
 * reveal which ids exist. Use it for a small closed team; plug an identity provider into
 * `authenticate` for anything larger.
 */
export function authenticateFromUsers(users) {
	const table = new Map(Object.entries(users).map(([id, entry]) => [id, accountOf(entry)]));
	let decoy;
	return async ({ user, credential }) => {
		const account = table.get(user);
		const expected = account?.credential;
		if (typeof credential !== 'string') return null;
		if (expected === undefined) {
			decoy ??= await hashPassword('decoy-credential');
			await verifyPassword(credential, decoy);
			return null;
		}
		if (typeof expected !== 'string') return null;
		const ok = isHashedCredential(expected)
			? await verifyPassword(credential, expected)
			: sameSecret(credential, expected);
		return ok ? { id: user, ...(account.role ? { role: account.role } : {}) } : null;
	};
}

/** User ids whose stored credential is plain text (for the startup warning). */
export const plainTextUsers = (users) =>
	Object.entries(users)
		.filter(([, value]) => !isHashedCredential(accountOf(value).credential))
		.map(([id]) => id);
