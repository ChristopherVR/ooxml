import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';

import { afterEach, describe, expect, it } from 'vitest';

import { aiSettings, streamReviewChat } from './ai-route';

const servers: ReturnType<typeof createServer>[] = [];
afterEach(() => {
	for (const server of servers.splice(0)) server.close();
});

describe('ai route', () => {
	it('needs a model, and a key unless the model is local', () => {
		expect(aiSettings({})).toBeNull();
		expect(aiSettings({ REVIEW_AI_MODEL: 'm' })).toBeNull();
		expect(aiSettings({ REVIEW_AI_MODEL: 'm', REVIEW_AI_API_KEY: 'k' })).toMatchObject({
			baseURL: 'https://api.openai.com/v1',
		});
		expect(
			aiSettings({ REVIEW_AI_MODEL: 'm', REVIEW_AI_BASE_URL: 'http://localhost:11434/v1' }),
		).toMatchObject({ apiKey: '', baseURL: 'http://localhost:11434/v1' });
	});

	it('streams a model reply with the deck tools offered to the model', async () => {
		let request: { tools?: unknown[]; messages?: { role: string }[] } = {};
		const model = createServer((incoming, response) => {
			let body = '';
			incoming.on('data', (chunk: Buffer) => (body += chunk));
			incoming.on('end', () => {
				request = JSON.parse(body);
				response.writeHead(200, { 'Content-Type': 'text/event-stream' });
				const chunk = (delta: object, finish: string | null) =>
					`data: ${JSON.stringify({ id: 'c', object: 'chat.completion.chunk', created: 0, model: 'm', choices: [{ index: 0, delta, finish_reason: finish }] })}\n\n`;
				response.end(
					`${chunk({ role: 'assistant', content: 'Looks good' }, null)}${chunk({}, 'stop')}data: [DONE]\n\n`,
				);
			});
		});
		servers.push(model);
		await new Promise<void>((resolve) => model.listen(0, '127.0.0.1', resolve));
		const baseURL = `http://127.0.0.1:${(model.address() as AddressInfo).port}/v1`;
		const reply = await streamReviewChat(
			{ baseURL, apiKey: '', model: 'm' },
			{ messages: [{ id: '1', role: 'user', parts: [{ type: 'text', text: 'Check slide 1' }] }] },
			{ reviewTitle: 'Q3 launch deck', reviewer: 'Ben Okafor' },
		);
		expect(reply.status).toBe(200);
		expect(await reply.text()).toContain('Looks good');
		expect(request.tools?.length).toBeGreaterThan(10);
		expect(request.messages?.map((message) => message.role)).toEqual(['system', 'user']);
	});

	it('rejects a body without messages', async () => {
		const reply = await streamReviewChat(
			{ baseURL: 'http://127.0.0.1:1/v1', apiKey: '', model: 'm' },
			{},
			{
				reviewTitle: 'x',
				reviewer: 'y',
			},
		);
		expect(reply.status).toBe(400);
	});
});
