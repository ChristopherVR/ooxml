import type {
	PptxAiBridge,
	PptxAiChatSession,
	PptxAiConfig,
	PptxAiUIMessage,
	ToolCanvasTarget,
} from 'ooxml-ui/pptx/ai';
import { describe, expect, it, vi } from 'vitest';

import { SvelteAiChat } from './chat.svelte';

/** Captures the options the panel passes to `@ai-sdk/svelte`'s `Chat`. */
interface CapturedChat {
	onToolCall?: (arg: { toolCall: unknown }) => Promise<void> | void;
	sendAutomaticallyWhen?: unknown;
	addToolOutput: ReturnType<typeof vi.fn>;
}
let captured: CapturedChat | undefined;

vi.mock(import('@ai-sdk/svelte'), () => ({
	Chat: class {
		messages: unknown[] = [];
		status = 'ready';
		error: undefined;
		addToolOutput = vi.fn();
		constructor(options: Pick<CapturedChat, 'onToolCall' | 'sendAutomaticallyWhen'>) {
			captured = { ...options, addToolOutput: this.addToolOutput };
		}
		sendMessage = vi.fn();
		stop = vi.fn();
		clearError = vi.fn();
	},
}));

function fakeBridge(): PptxAiBridge {
	return {
		getDeckMeta: () => ({ slideCount: 1, activeSlideIndex: 0, width: 960, height: 540 }),
		getSlides: () => [],
		getActiveSlideIndex: () => 0,
		getTheme: () => undefined,
		getHandler: () => undefined,
		getDeckLoad: () => 1,
		goToSlide: () => undefined,
		selectElements: () => undefined,
		applySlidesUpdate: () => undefined,
		updateElement: () => undefined,
		applyTheme: () => undefined,
	};
}

function fakeSession(): PptxAiChatSession {
	return {
		transport: {} as never,
		sendAutomaticallyWhen: () => false,
		clientExecutesTools: true,
		executeToolCall: vi.fn().mockResolvedValue({ ok: true }),
		proposals: {
			list: () => [],
			apply: vi.fn(),
			revert: vi.fn(),
			acceptAll: vi.fn(),
			clear: vi.fn(),
		},
	} as unknown as PptxAiChatSession;
}

const config: PptxAiConfig = { connection: { kind: 'transport', transport: {} as never } };

function toolMessage(toolCallId: string, input: unknown): PptxAiUIMessage {
	return {
		id: `m-${toolCallId}`,
		role: 'assistant',
		parts: [{ type: 'tool-update_element', toolCallId, state: 'input-available', input }],
	} as unknown as PptxAiUIMessage;
}

async function readyChat(clientExecutesTools = true) {
	captured = undefined;
	const targets: (ToolCanvasTarget | null)[] = [];
	const chat = new SvelteAiChat({
		bridge: fakeBridge(),
		config,
		checkAvailable: () => Promise.resolve(true),
		createSession: () => Promise.resolve({ ...fakeSession(), clientExecutesTools }),
		onToolTarget: (target) => targets.push(target),
	});
	await chat.init();
	return { chat, targets };
}

describe('svelteAiChat tool loop', () => {
	it('runs a tool the client is responsible for and returns its output', async () => {
		await readyChat();
		await captured?.onToolCall?.({
			toolCall: { toolName: 'update_element', toolCallId: 'c1', input: {} },
		});
		expect(captured?.addToolOutput).toHaveBeenCalledWith(
			expect.objectContaining({ tool: 'update_element', toolCallId: 'c1' }),
		);
	});

	it('leaves the tools to the agent in model mode, so each runs once', async () => {
		await readyChat(false);
		expect(captured?.onToolCall).toBeUndefined();
		expect(captured?.sendAutomaticallyWhen).toBeUndefined();
	});
});

describe('svelteAiChat live canvas focus', () => {
	it('reports each tool call in the stream once, in every connection mode', async () => {
		const { chat, targets } = await readyChat(false);
		chat.messages.push(toolMessage('c1', { slideIndex: 4, elementId: 'shape-9', text: 'Hi' }));

		chat.followToolCalls();
		chat.followToolCalls();

		expect(targets).toStrictEqual([{ slideIndex: 4, elementIds: ['shape-9'] }]);
	});

	it('lists the proposals the agent staged while the reply streams', async () => {
		const staged = [{ id: 'p1', label: 'Edit', summary: [], createdAt: 1 }];
		captured = undefined;
		const session = fakeSession();
		session.proposals.list = () => staged as ReturnType<typeof session.proposals.list>;
		const chat = new SvelteAiChat({
			bridge: fakeBridge(),
			config,
			checkAvailable: () => Promise.resolve(true),
			createSession: () => Promise.resolve({ ...session, clientExecutesTools: false }),
		});
		await chat.init();

		chat.followToolCalls();

		expect(chat.proposals).toEqual(staged);
	});

	it('does not replay the tool calls of a resumed chat', async () => {
		const { chat, targets } = await readyChat();
		chat.setMessages([toolMessage('old', { slideIndex: 1, elementId: 'e', text: 'x' })]);
		chat.followToolCalls();
		expect(targets).toStrictEqual([]);
	});
});
