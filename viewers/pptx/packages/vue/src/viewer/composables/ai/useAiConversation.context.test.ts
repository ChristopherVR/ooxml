import { useChat } from '@ai-sdk/vue';
import { mount } from '@vue/test-utils';
import type { ChatTransport } from 'ai';
import { createAiChatSession, ProposalStore } from 'ooxml-ui/pptx/ai';
import type {
	PptxAiBridge,
	PptxAiChatSession,
	PptxAiConfig,
	PptxAiUIMessage,
} from 'ooxml-ui/pptx/ai';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import { defineComponent, h, nextTick, ref } from 'vue';

import { useAiConversation } from './useAiConversation';

const slides = [
	{ id: 's1', slideNumber: 1, elements: [] },
	{ id: 's2', slideNumber: 2, elements: [] },
] as unknown as PptxSlide[];

const bridge = {
	getDeckMeta: () => ({ slideCount: 2, activeSlideIndex: 0, width: 960, height: 540 }),
	getSlides: () => slides,
	getActiveSlideIndex: () => 0,
	getTheme: () => undefined,
	getHandler: () => undefined,
	getDeckLoad: () => 0,
	goToSlide: () => undefined,
	selectElements: () => undefined,
	applySlidesUpdate: () => undefined,
	updateElement: () => undefined,
	applyTheme: () => undefined,
} satisfies PptxAiBridge;

function mountWithChatStub(session: PptxAiChatSession, deckLoad: () => number = () => 0) {
	let options: Record<string, unknown> = {};
	const messages = ref<PptxAiUIMessage[]>([]);
	const useChatStub = ((init: Record<string, unknown>) => {
		options = init;
		return { messages, status: ref('ready'), error: ref(undefined) };
	}) as unknown as typeof useChat;
	let conversation: ReturnType<typeof useAiConversation> | undefined;
	mount(
		defineComponent({
			setup() {
				conversation = useAiConversation(useChatStub, session, {} as PptxAiConfig, { deckLoad });
				return () => h('div');
			},
		}),
	);
	return {
		options,
		messages,
		proposals: () => conversation?.proposals.value,
		discarded: () => conversation?.proposalsDiscarded.value,
	};
}

function stubSession(clientExecutesTools: boolean): PptxAiChatSession {
	return {
		transport: {},
		clientExecutesTools,
		sendAutomaticallyWhen: () => true,
		executeToolCall: vi.fn(),
		proposals: { list: () => [] },
	} as unknown as PptxAiChatSession;
}

describe('useAiConversation', () => {
	it('leaves the tools to the agent in model mode, so each runs once', () => {
		const { options } = mountWithChatStub(stubSession(false));
		expect(options.onToolCall).toBeUndefined();
		expect(options.sendAutomaticallyWhen).toBeUndefined();
	});

	it('runs the tools itself for an endpoint connection', () => {
		expect(mountWithChatStub(stubSession(true)).options.onToolCall).toBeTypeOf('function');
	});

	it('lists a proposal the agent staged while the reply streams', async () => {
		const session = stubSession(false);
		const staged = [{ id: 'p1', label: 'Edit', summary: [], createdAt: 1 }];
		const { messages, proposals } = mountWithChatStub(session);
		expect(proposals()).toEqual([]);

		session.proposals.list = () => staged as ReturnType<typeof session.proposals.list>;
		messages.value = [{ id: 'm1', role: 'assistant', parts: [] }] as PptxAiUIMessage[];
		await nextTick();

		expect(proposals()).toEqual(staged);
	});

	it('drops the listed proposals as soon as another deck loads, and says so', async () => {
		const deckLoad = ref(0);
		const proposals = new ProposalStore({ ...bridge, getDeckLoad: () => deckLoad.value });
		proposals.stage('Edit', (deck) => deck.slice(1));
		const session = { ...stubSession(false), proposals } as PptxAiChatSession;
		const conversation = mountWithChatStub(session, () => deckLoad.value);
		expect(conversation.proposals()).toHaveLength(1);
		expect(conversation.discarded()).toBeFalsy();

		deckLoad.value += 1;
		await nextTick();

		expect(conversation.proposals()).toStrictEqual([]);
		expect(conversation.discarded()).toBeTruthy();
	});

	it('sends the deck context the session adds to every request', async () => {
		const requests: PptxAiUIMessage[][] = [];
		const transport = {
			async sendMessages(options: { messages: PptxAiUIMessage[] }) {
				requests.push(options.messages);
				return new ReadableStream({
					start(controller) {
						controller.enqueue({ type: 'start' });
						controller.enqueue({ type: 'finish', finishReason: 'stop' });
						controller.close();
					},
				});
			},
			reconnectToStream: async () => null,
		} as unknown as ChatTransport<PptxAiUIMessage>;
		const config: PptxAiConfig = { connection: { kind: 'transport', transport } };
		const session = await createAiChatSession(bridge, config);
		let send: (text: string) => void = () => undefined;
		mount(
			defineComponent({
				setup() {
					send = useAiConversation(useChat, session, config, { deckLoad: () => 0 }).send;
					return () => h('div');
				},
			}),
		);

		send('hello');

		await vi.waitFor(() => expect(requests).toHaveLength(1));
		expect(JSON.stringify(requests[0])).toContain('Deck: Untitled - 2 slide(s)');
	});
});
