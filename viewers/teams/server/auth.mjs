// Session tokens for the reference server. A client proves who it is once (POST /auth/session,
// through the `authenticate` hook or the built-in `users` table) and gets a signed, expiring token
// that binds one user id. Sockets then present that token, and the server uses the id inside it,
// never the `user` query parameter. Tokens are HMAC-SHA-256 over the payload; the secret never
// leaves the process and is never logged. Pure helpers, no I/O.
import { createHash, createHmac, randomBytes, timingSafeEqual } from 'node:crypto';

const USER = /^[\w.@:+-]{1,64}$/u;
export const DEFAULT_SESSION_TTL_MS = 12 * 60 * 60 * 1000;
const b64 = (buffer) => Buffer.from(buffer).toString('base64url');

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
 */
export function createSessionSigner({
	secret,
	ttlMs = DEFAULT_SESSION_TTL_MS,
	now = Date.now,
} = {}) {
	const key = createHash('sha256')
		.update(`teams-session|${secret || randomBytes(32).toString('hex')}`)
		.digest();
	const mac = (payload) => createHmac('sha256', key).update(payload).digest();
	return {
		ttlMs,
		/** A token for `user`, valid for `ttlMs`. */
		issue(user) {
			if (!isUserId(user)) throw new Error('invalid user id');
			const exp = now() + ttlMs;
			const payload = b64(JSON.stringify({ u: user, exp }));
			return { token: `v1.${payload}.${b64(mac(payload))}`, user, expiresAt: exp };
		},
		/** The authenticated user id of a token, or null when it is forged, malformed or expired. */
		verify(token) {
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
				return claims.u;
			} catch {
				return null;
			}
		},
	};
}

/**
 * The credential table behind `users`: user id -> credential. Returns an `authenticate` function.
 * Credentials are compared in constant time. Use it for a small closed team; plug an identity
 * provider into `authenticate` for anything larger.
 */
export function authenticateFromUsers(users) {
	const table = new Map(Object.entries(users));
	return ({ user, credential }) => {
		const expected = table.get(user);
		return expected !== undefined &&
			typeof credential === 'string' &&
			sameSecret(credential, expected)
			? { id: user }
			: null;
	};
}
