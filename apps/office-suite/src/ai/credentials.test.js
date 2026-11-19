import assert from 'node:assert/strict';
import { afterEach, test } from 'node:test';

import { installStorage } from './test-storage.js';

installStorage();
const { roomToken } = await import('./credentials.js');

const { fetch } = globalThis;
afterEach(() => {
	globalThis.fetch = fetch;
});

/** Answers each credentials request with the next of `bodies`, recording what was asked. */
function issuer(...bodies) {
	const calls = [];
	globalThis.fetch = async (url, init) => {
		calls.push({ url, init });
		const body = bodies.shift();
		return body instanceof Response ? body : Response.json(body);
	};
	return calls;
}
const inAnHour = () => new Date(Date.now() + 3_600_000).toISOString();

test('the token is fetched with the user cookies, held in memory and refetched on reconnect', async () => {
	const calls = issuer(
		{ token: 'tok-1', expiresAt: inAnHour() },
		{ token: 'tok-2', expiresAt: inAnHour() },
	);
	const url = 'https://host.example/a';
	const first = roomToken(url);
	assert.equal(await first(), 'tok-1');
	assert.equal(calls[0].url, url);
	assert.equal(calls[0].init.credentials, 'include');
	// Another room connecting while the token is valid reuses it.
	assert.equal(await roomToken(url)(), 'tok-1');
	assert.equal(calls.length, 1);
	// A reconnect asks again.
	assert.equal(await first(), 'tok-2');
	assert.equal(calls.length, 2);
	assert.doesNotMatch(localStorage.dump() + sessionStorage.dump(), /tok-/u);
});

test('an expired token is fetched again', async () => {
	const calls = issuer(
		{ token: 'short', expiresAt: new Date(Date.now() + 1000).toISOString() },
		{ token: 'next', expiresAt: inAnHour() },
	);
	const url = 'https://host.example/b';
	assert.equal(await roomToken(url)(), 'short');
	assert.equal(await roomToken(url)(), 'next');
	assert.equal(calls.length, 2);
});

test('a refused or malformed answer rejects with a reason the chat can show', async () => {
	issuer(new Response('', { status: 401 }), { token: '', expiresAt: inAnHour() });
	await assert.rejects(roomToken('https://host.example/c')(), /not syncing.*HTTP 401/u);
	await assert.rejects(roomToken('https://host.example/c')(), /no valid credential/u);
});
