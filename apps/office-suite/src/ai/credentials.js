/**
 * Short-lived sync credentials for shared chats. When the policy names a `credentialsUrl`, the
 * suite asks that host endpoint for a token right before a shared chat connects, with the user's
 * cookies (`credentials: 'include'`), and expects `{ "token": "...", "expiresAt": "<ISO time>" }`.
 * The token is held in memory only, never in storage, and is fetched again once it expires and
 * before every reconnect.
 */

/** How long before its `expiresAt` a held token counts as expired. */
const MARGIN_MS = 30_000;
let held = null;

async function issue(url) {
	let response;
	try {
		response = await fetch(url, {
			credentials: 'include',
			cache: 'no-store',
			redirect: 'error',
			headers: { Accept: 'application/json' },
		});
	} catch {
		throw new Error('Shared chats are not syncing: the sign-in service could not be reached.');
	}
	if (!response.ok)
		throw new Error(
			`Shared chats are not syncing: the sign-in service returned HTTP ${response.status}.`,
		);
	let body;
	try {
		body = await response.json();
	} catch {}
	const expires = Date.parse(body?.expiresAt);
	if (
		typeof body?.token !== 'string' ||
		!body.token ||
		body.token.length > 4096 ||
		!(expires > Date.now())
	)
		throw new Error('Shared chats are not syncing: the sign-in service sent no valid credential.');
	return { url, token: body.token, expires };
}

/**
 * The token for the shared-chat server: the one held, unless `fresh` or it expired, else a new
 * one from `url`. Rejects with a sentence the chat can show.
 */
export async function syncToken(url, fresh = false) {
	if (fresh || held?.url !== url || held.expires - MARGIN_MS <= Date.now()) held = await issue(url);
	return held.token;
}

/** A token source for one room: the held token on the first connection, a new one on reconnects. */
export function roomToken(url) {
	let connected = false;
	return () => {
		const fresh = connected;
		connected = true;
		return syncToken(url, fresh);
	};
}
