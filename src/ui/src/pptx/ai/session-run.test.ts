import type { ChatTransport, UIMessageChunk } from 'ai';
import { describe, expect, it } from 'vitest';

import type { PptxAiConfig, PptxAiUIMessage } from './config';
import { makeMockBridge } from './mock-bridge';
import type { PptxAiRunLimits } from './run-limits';
import { RunLimiter } from './run-limits';
import { createAiChatSession } from './session';
import { createSessionTransport, injectDeckContext } from './session-transport';
import type { VanillaChatController } from './vanilla-chat';
import { createVanillaChat } from './vanilla-chat';

const streamOf = (chunks: UIMessageChunk[]) =>
	new ReadableStream<UIMessageChunk>({
		start(controller) {
			for (const chunk of chunks) {
				controller.enqueue(chunk);
			}
			controller.close();
		},
	});

const toolStep = (id: string, toolName: string, input: unknown): UIMessageChunk[] => [
	{ type: 'start' },
	{ type: 'start-step' },
	{ type: 'tool-input-start', toolCallId: id, toolName },
	{ type: 'tool-input-available', toolCallId: id, toolName, input },
	{ type: 'finish-step' },
	{ type: 'finish', finishReason: 'tool-calls' },
];

const textStep = (text: string, finishReason?: 'stop' | 'length'): UIMessageChunk[] => [
	{ type: 'start' },
	{ type: 'start-step' },
	{ type: 'text-start', id: 't' },
	{ type: 'text-delta', id: 't', delta: text },
	{ type: 'text-end', id: 't' },
	{ type: 'finish-step' },
	{ type: 'finish', ...(finishReason ? { finishReason } : {}) },
];

type Requests = Parameters<ChatTransport<PptxAiUIMessage>['sendMessages']>[0][];

function scriptedTransport(next: (call: number) => UIMessageChunk[]) {
	const requests: Requests = [];
	const transport: ChatTransport<PptxAiUIMessage> = {
		async sendMessages(options) {
			requests.push(options);
			return streamOf(next(requests.length));
		},
		async reconnectToStream() {
			return null;
		},
	};
	return { transport, requests };
}

async function settled(chat: VanillaChatController): Promise<void> {
	await chat.sendMessage('go');
	const start = Date.now();
	for (;;) {
		const { status } = chat.getSnapshot();
		if (status === 'ready' || status === 'error') {
			await new Promise((resolve) => setTimeout(resolve, 5));
			if (chat.getSnapshot().status === status) {
				return;
			}
		}
		if (Date.now() - start > 4000) {
			throw new Error('The run did not end');
		}
		await new Promise((resolve) => setTimeout(resolve, 5));
	}
}

async function runScripted(next: (call: number) => UIMessageChunk[], limits?: PptxAiRunLimits) {
	const { transport, requests } = scriptedTransport(next);
	const bridge = makeMockBridge();
	const config: PptxAiConfig = {
		connection: { kind: 'transport', transport },
		...(limits ? { limits } : {}),
	};
	const chat = await createVanillaChat({ bridge, config });
	await settled(chat);
	return { chat, requests, bridge };
}

describe('run limits', () => {
	it('stops a model that keeps repeating one call', async () => {
		const { chat, requests } = await runScripted((n) => toolStep(`c${n}`, 'get_deck_overview', {}));

		expect(requests.length).toBe(5);
		expect(chat.getLastFinish()).toEqual({ reason: 'limit', limit: 'repeats' });
		expect(chat.getSnapshot().status).toBe('error');
		expect(chat.getSnapshot().error?.message).toContain('repeating');
	});

	it('stops at the step limit', async () => {
		const { chat, requests } = await runScripted(
			(n) => toolStep(`c${n}`, 'get_slide', { slideIndex: n % 2 }),
			{ maxSteps: 3, maxRepeatedToolCalls: 100 },
		);

		expect(requests.length).toBe(3);
		expect(chat.getLastFinish()).toEqual({ reason: 'limit', limit: 'steps' });
	});

	it('stops at the tool call limit', async () => {
		const { chat, requests } = await runScripted(
			(n) => toolStep(`c${n}`, 'get_slide', { slideIndex: n % 2 }),
			{ maxToolCalls: 2, maxRepeatedToolCalls: 100 },
		);

		expect(requests.length).toBe(3);
		expect(chat.getLastFinish()).toEqual({ reason: 'limit', limit: 'tool-calls' });
	});

	it('does not run the tool call that goes over a limit', async () => {
		const { bridge } = await runScripted(
			(n) => toolStep(`c${n}`, 'go_to_slide', { slideIndex: 1 }),
			{ maxRepeatedToolCalls: 2 },
		);

		expect(bridge.navigations).toEqual([1, 1]);
	});

	it('stops a request that outlives the time limit', async () => {
		let requests = 0;
		const transport: ChatTransport<PptxAiUIMessage> = {
			async sendMessages() {
				requests += 1;
				return new ReadableStream<UIMessageChunk>({
					start(controller) {
						controller.enqueue({ type: 'start' });
					},
				});
			},
			async reconnectToStream() {
				return null;
			},
		};
		const chat = await createVanillaChat({
			bridge: makeMockBridge(),
			config: { connection: { kind: 'transport', transport }, limits: { maxRunMs: 30 } },
		});

		await settled(chat);

		expect(requests).toBe(1);
		expect(chat.getLastFinish()).toEqual({ reason: 'limit', limit: 'time' });
	});

	it('starts counting again for the next user message', async () => {
		const limiter = new RunLimiter({ maxSteps: 1 });
		limiter.begin();
		expect(limiter.admitStep()).toBe(true);
		expect(limiter.admitStep()).toBe(false);

		limiter.begin();

		expect(limiter.tripped).toBeUndefined();
		expect(limiter.admitStep()).toBe(true);
	});
});

describe('finish reasons', () => {
	it('reports a normal end', async () => {
		const { chat } = await runScripted(() => textStep('Done', 'stop'));

		expect(chat.getLastFinish()).toEqual({ reason: 'stop' });
		expect(chat.getSnapshot().status).toBe('ready');
	});

	it('tells the user when a reply was cut off, and keeps the text', async () => {
		const { chat } = await runScripted(() => textStep('Partial', 'length'));

		expect(chat.getLastFinish()).toEqual({ reason: 'length' });
		expect(chat.getSnapshot().error?.message).toContain('cut off');
		expect(JSON.stringify(chat.getSnapshot().messages)).toContain('Partial');
	});

	it('treats a stream with no finish as an unknown ending', async () => {
		const { chat } = await runScripted(() => textStep('Half').slice(0, 4));

		expect(chat.getLastFinish()).toEqual({ reason: 'unknown' });
		expect(chat.getSnapshot().status).toBe('error');
	});
});

describe('deck context', () => {
	const textOf = (message: PptxAiUIMessage | undefined) => JSON.stringify(message?.parts ?? []);

	it('reaches the model on every request and stays out of the transcript', async () => {
		const { chat, requests } = await runScripted((n) =>
			n === 1 ? toolStep('c1', 'get_deck_overview', {}) : textStep('Done', 'stop'),
		);

		expect(requests).toHaveLength(2);
		for (const request of requests) {
			expect(textOf(request.messages[0])).toContain('Deck: Untitled - 2 slide(s)');
		}
		expect(textOf(chat.getSnapshot().messages[0])).not.toContain('Deck:');
	});

	it('includes what the user focused', async () => {
		const { transport, requests } = scriptedTransport(() => textStep('Done', 'stop'));
		const bridge = makeMockBridge();
		bridge.getFocusedTargets = () => [{ kind: 'element', slideIndex: 0, elementId: 'el-2' }];
		const chat = await createVanillaChat({
			bridge,
			config: { connection: { kind: 'transport', transport }, contextStrategy: 'none' },
		});

		await settled(chat);

		expect(textOf(requests[0]?.messages[0])).toContain('text#el-2');
		expect(textOf(requests[0]?.messages[0])).not.toContain('Deck:');
	});

	it('sends nothing extra when context is off and nothing is focused', async () => {
		const { transport, requests } = scriptedTransport(() => textStep('Done', 'stop'));
		const wrapped = createSessionTransport(transport, {
			bridge: makeMockBridge(),
			contextStrategy: 'none',
			limiter: new RunLimiter(),
			onFinish() {},
		});
		const messages = [
			{ id: 'u', role: 'user', parts: [{ type: 'text', text: 'hi' }] },
		] as PptxAiUIMessage[];

		await wrapped.sendMessages({
			trigger: 'submit-message',
			chatId: 'c',
			messageId: undefined,
			messages,
			abortSignal: undefined,
		});

		expect(requests[0]?.messages).toEqual(messages);
	});

	it('adds the block to the last user message without mutating the input', () => {
		const messages = [
			{ id: '1', role: 'user', parts: [{ type: 'text', text: 'first' }] },
			{ id: '2', role: 'assistant', parts: [{ type: 'text', text: 'ok' }] },
			{ id: '3', role: 'user', parts: [{ type: 'text', text: 'second' }] },
		] as PptxAiUIMessage[];

		const next = injectDeckContext(messages, 'CTX');

		expect(next[2]?.parts).toHaveLength(2);
		expect(JSON.stringify(next[2]?.parts[0])).toContain('CTX');
		expect(next[0]).toBe(messages[0]);
		expect(messages[2]?.parts).toHaveLength(1);
	});

	it('is also applied to a session used directly by a binding', async () => {
		const { transport, requests } = scriptedTransport(() => textStep('Done', 'stop'));
		const session = await createAiChatSession(makeMockBridge(), {
			connection: { kind: 'transport', transport },
		});

		await session.transport.sendMessages({
			trigger: 'submit-message',
			chatId: 'c',
			messageId: undefined,
			messages: [{ id: 'u', role: 'user', parts: [{ type: 'text', text: 'hi' }] }],
			abortSignal: undefined,
		});

		expect(JSON.stringify(requests[0]?.messages)).toContain('Deck: Untitled');
	});
});
