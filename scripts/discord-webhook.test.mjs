import assert from 'node:assert/strict';
import { test } from 'node:test';

import { hookBody, KINDS, parseWebhook } from './discord-webhook.mjs';

const URL_OK = 'https://discord.com/api/webhooks/123456789/abc_DEF-123';

test('parseWebhook accepts a Discord webhook URL and strips a trailing /github', () => {
	assert.equal(parseWebhook(`  ${URL_OK}\r\n`), URL_OK);
	assert.equal(parseWebhook(`${URL_OK}/github`), URL_OK);
});

test('parseWebhook rejects anything else', () => {
	assert.equal(parseWebhook('hello'), null);
	assert.equal(parseWebhook('https://example.com/api/webhooks/1/x'), null);
	assert.equal(parseWebhook(`${URL_OK}?x=1`), null);
});

test('hookBody points at the native GitHub endpoint with the requested events', () => {
	const body = hookBody(URL_OK, KINDS.prs.events);
	assert.equal(body.config.url, `${URL_OK}/github`);
	assert.equal(body.config.content_type, 'json');
	assert.deepEqual(body.events, KINDS.prs.events);
});

test('every kind is either a webhook or a secret, and pr and issue events do not overlap', () => {
	for (const spec of Object.values(KINDS)) assert.ok(Boolean(spec.events) !== Boolean(spec.secret));
	const overlap = KINDS.prs.events.filter((e) => KINDS.issues.events.includes(e));
	assert.deepEqual(overlap, []);
});
