// Password hashing for the `users` table. Format: `scrypt$N$r$p$<salt b64url>$<hash b64url>`, made
// by node:crypto scrypt with a random 16 byte salt. Verification is constant-time and async, so a
// login does not block the event loop. A stored value without the `scrypt$` prefix is a plain text
// credential (still accepted, see `isHashedCredential`). Passwords are never logged.
import { randomBytes, scrypt, timingSafeEqual } from 'node:crypto';

const PREFIX = 'scrypt$';
const KEY_BYTES = 32;
const DEFAULTS = { N: 16384, r: 8, p: 1 };

const derive = (password, salt, { N, r, p }) =>
	new Promise((resolve, reject) =>
		scrypt(
			password.normalize('NFKC'),
			salt,
			KEY_BYTES,
			{ N, r, p, maxmem: 128 * N * r + 1024 * 1024 },
			(error, key) => (error ? reject(error) : resolve(key)),
		),
	);

export const isHashedCredential = (stored) =>
	typeof stored === 'string' && stored.startsWith(PREFIX);

/** Hash `password` for the users table. */
export async function hashPassword(password, params = DEFAULTS) {
	if (typeof password !== 'string' || password.length === 0) throw new Error('empty password');
	const salt = randomBytes(16);
	const key = await derive(password, salt, params);
	return `${PREFIX}${params.N}$${params.r}$${params.p}$${salt.toString('base64url')}$${key.toString('base64url')}`;
}

/** True when `password` matches `stored` (a hash made by `hashPassword`). Never throws. */
export async function verifyPassword(password, stored) {
	try {
		if (typeof password !== 'string' || !isHashedCredential(stored)) return false;
		const [, n, r, p, salt, hash] = stored.split('$');
		const params = { N: Number(n), r: Number(r), p: Number(p) };
		// Refuse absurd cost parameters from a corrupt file instead of exhausting memory.
		if (!(params.N >= 2 && params.N <= 1 << 20 && (params.N & (params.N - 1)) === 0)) return false;
		if (!(params.r >= 1 && params.r <= 32 && params.p >= 1 && params.p <= 16)) return false;
		const expected = Buffer.from(hash ?? '', 'base64url');
		const given = await derive(password, Buffer.from(salt ?? '', 'base64url'), params);
		return expected.length === given.length && timingSafeEqual(expected, given);
	} catch {
		return false;
	}
}
