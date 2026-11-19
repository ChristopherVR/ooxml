import assert from 'node:assert/strict';
import { afterEach, beforeEach, test } from 'node:test';

const storage = new Map();
globalThis.localStorage ??= {
	getItem: (key) => storage.get(key) ?? null,
	setItem: (key, value) => storage.set(key, value),
	removeItem: (key) => storage.delete(key),
};
globalThis.sessionStorage ??= { getItem: () => null, setItem() {}, removeItem() {} };

const { normalizeFinishReason, ReplyEndedError, streamChat } = await import('./providers.js');

const openai = { id: 'o', provider: 'openai', endpoint: 'https://x.test/v1', model: 'm', key: 'k' };
const anthropic = {
	id: 'a',
	provider: 'anthropic',
	endpoint: 'https://x.test/m',
	model: 'm',
	key: 'k',
};

const sse = (events) =>
	new Response(
		new ReadableStream({
			start(controller) {
				const encoder = new TextEncoder();
				for (const event of events)
					controller.enqueue(
						encoder.encode(
							`data: ${typeof event === 'string' ? event : JSON.stringify(event)}\n\n`,
						),
					);
				controller.close();
			},
		}),
		{ headers: { 'content-type': 'text/event-stream' } },
	);

const text = (content) => ({ choices: [{ delta: { content } }] });
const end = (finish_reason) => ({ choices: [{ delta: {}, finish_reason }] });
const anthropicText = (value) => ({
	type: 'content_block_delta',
	delta: { type: 'text_delta', text: value },
});
const anthropicEnd = (stop_reason) => ({ type: 'message_delta', delta: { stop_reason } });

const realFetch = globalThis.fetch;
let reply;
beforeEach(() => {
	globalThis.fetch = async () => reply;
});
afterEach(() => {
	globalThis.fetch = realFetch;
});

function send(connection, response) {
	reply = response;
	return streamChat({
		connection,
		system: 'S',
		messages: [{ role: 'user', content: 'hi' }],
		signal: new AbortController().signal,
		onText() {},
		context: null,
		settings: { maxTokens: 10 },
	});
}

const ended = (reason, partial) => (error) => {
	assert.ok(error instanceof ReplyEndedError);
	assert.equal(error.reason, reason);
	assert.equal(error.partial, partial);
	return true;
};

test('a reply that finishes normally is returned', async () => {
	assert.equal(await send(openai, sse([text('Done'), end('stop'), '[DONE]'])), 'Done');
	assert.equal(await send(openai, sse([text('Done'), '[DONE]'])), 'Done');
	assert.equal(
		await send(
			anthropic,
			sse([anthropicText('Done'), anthropicEnd('end_turn'), { type: 'message_stop' }]),
		),
		'Done',
	);
});

test('a reply cut off by the length limit keeps its text and says so', async () => {
	await assert.rejects(
		send(openai, sse([text('Par'), end('length'), '[DONE]'])),
		ended('length', 'Par'),
	);
	await assert.rejects(
		send(
			anthropic,
			sse([anthropicText('Par'), anthropicEnd('max_tokens'), { type: 'message_stop' }]),
		),
		ended('length', 'Par'),
	);
});

test('a filtered reply is reported', async () => {
	await assert.rejects(
		send(openai, sse([text('Hel'), end('content_filter')])),
		ended('filtered', 'Hel'),
	);
});

test('a stream with no finish signal is an error, not a complete reply', async () => {
	await assert.rejects(send(openai, sse([text('Half a sen')])), ended('unknown', 'Half a sen'));
	await assert.rejects(send(anthropic, sse([anthropicText('Half')])), ended('unknown', 'Half'));
});

test('a JSON reply reports a stated ending', async () => {
	const body = (finish_reason) =>
		new Response(JSON.stringify({ choices: [{ message: { content: 'Par' }, finish_reason }] }), {
			headers: { 'content-type': 'application/json' },
		});
	assert.equal(await send(openai, body('stop')), 'Par');
	await assert.rejects(send(openai, body('length')), ended('length', 'Par'));
});

test('an unrecognised ending normalises to unknown', () => {
	assert.equal(normalizeFinishReason('something_new'), 'unknown');
	assert.equal(normalizeFinishReason('end_turn'), 'stop');
});
