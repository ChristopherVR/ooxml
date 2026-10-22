// Signed, expiring tokens: the host's sign-in cookie and the short-lived room tickets. The browser
// fetches a ticket with its ordinary host session and hands it to the collaboration socket; the
// relay trusts only what the ticket says (account, review and access), never the client.
import { createHmac, timingSafeEqual } from 'node:crypto';

import type { Permission } from './directory';

export interface Ticket {
	account: string;
	review: string;
	permission: Permission;
	/** Expiry in epoch milliseconds. A ticket only admits a new connection; it does not end one. */
	exp: number;
}

export interface SignIn {
	account: string;
	exp: number;
}

export const TICKET_TTL_MS = 5 * 60 * 1000;
export const SIGN_IN_TTL_MS = 8 * 60 * 60 * 1000;

const sign = (secret: string, body: string): string =>
	createHmac('sha256', secret).update(body).digest('base64url');

export function signToken(secret: string, payload: Ticket | SignIn): string {
	const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
	return `${body}.${sign(secret, body)}`;
}

/** The payload of a valid, unexpired token, or null. Callers still check its fields. */
export function verifyToken(
	secret: string,
	token: string,
	now = Date.now(),
): Record<string, unknown> | null {
	const [body, signature, extra] = token.split('.');
	if (!body || !signature || extra !== undefined) return null;
	const expected = Buffer.from(sign(secret, body));
	const actual = Buffer.from(signature);
	if (expected.length !== actual.length || !timingSafeEqual(expected, actual)) return null;
	try {
		const payload: unknown = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
		if (!payload || typeof payload !== 'object') return null;
		const record = payload as Record<string, unknown>;
		return typeof record.exp === 'number' && record.exp > now ? record : null;
	} catch {
		return null;
	}
}

const PERMISSIONS = new Set<unknown>(['owner', 'review', 'view']);

export function verifyTicket(secret: string, token: string, now = Date.now()): Ticket | null {
	const payload = verifyToken(secret, token, now);
	if (!payload || typeof payload.account !== 'string' || typeof payload.review !== 'string')
		return null;
	if (!PERMISSIONS.has(payload.permission)) return null;
	return payload as unknown as Ticket;
}

export function verifySignIn(secret: string, token: string, now = Date.now()): SignIn | null {
	const payload = verifyToken(secret, token, now);
	return payload && typeof payload.account === 'string' && payload.review === undefined
		? (payload as unknown as SignIn)
		: null;
}
