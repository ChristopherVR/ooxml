import type { ChatTransport } from 'ai';
import type { PptxAiBridge, PptxAiUIMessage } from 'ooxml-ui/pptx/ai';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createAiPanel } from './ai-panel';

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

type Chunk = Record<string, unknown>;

function scriptedTransport(next: (call: number) => Chunk[]) {
	const requests: PptxAiUIMessage[][] = [];
	const transport = {
		async sendMessages(options: { messages: PptxAiUIMessage[] }) {
			requests.push(options.messages);
			const chunks = next(requests.length);
			return new ReadableStream({
				start(controller) {
					for (const chunk of chunks) {
						controller.enqueue(chunk);
					}
					controller.close();
				},
			});
		},
		reconnectToStream: async () => null,
	} as unknown as ChatTransport<PptxAiUIMessage>;
	return { transport, requests };
}

let host: HTMLElement | null = null;
afterEach(() => {
	host?.remove();
	host = null;
});

async function openPanel(transport: ChatTransport<PptxAiUIMessage>): Promise<HTMLElement> {
	const element = document.createElement('div');
	host = element;
	document.body.appendChild(element);
	await createAiPanel({
		host: element,
		doc: document,
		t: createTranslator('en'),
		bridge,
		config: { connection: { kind: 'transport', transport } },
	});
	const input = element.querySelector('textarea') as HTMLTextAreaElement;
	input.value = 'hello';
	element
		.querySelector('form')
		?.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
	return element;
}

describe('createAiPanel over a real session', () => {
	it('sends the deck context the session adds to every request', async () => {
		const { transport, requests } = scriptedTransport(() => [
			{ type: 'start' },
			{ type: 'finish', finishReason: 'stop' },
		]);

		await openPanel(transport);

		await vi.waitFor(() => expect(requests).toHaveLength(1));
		expect(JSON.stringify(requests[0])).toContain('Deck: Untitled - 2 slide(s)');
	});

	it('shows why a run stopped at a limit', async () => {
		const { transport, requests } = scriptedTransport((call) => [
			{ type: 'start' },
			{ type: 'start-step' },
			{ type: 'tool-input-start', toolCallId: `c${call}`, toolName: 'get_deck_overview' },
			{
				type: 'tool-input-available',
				toolCallId: `c${call}`,
				toolName: 'get_deck_overview',
				input: {},
			},
			{ type: 'finish-step' },
			{ type: 'finish', finishReason: 'tool-calls' },
		]);

		const panel = await openPanel(transport);

		await vi.waitFor(() =>
			expect(panel.querySelector('.pptxv-ai-error')?.textContent).toContain('repeating'),
		);
		expect(requests).toHaveLength(5);
	});
});
