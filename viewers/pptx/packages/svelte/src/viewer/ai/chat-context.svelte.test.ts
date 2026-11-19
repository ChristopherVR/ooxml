import type { ChatTransport } from 'ai';
import type { PptxAiBridge, PptxAiUIMessage } from 'ooxml-ui/pptx/ai';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { SvelteAiChat } from './chat.svelte';

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
	getDeckLoad: () => 1,
	goToSlide: () => undefined,
	selectElements: () => undefined,
	applySlidesUpdate: () => undefined,
	updateElement: () => undefined,
	applyTheme: () => undefined,
} satisfies PptxAiBridge;

describe('svelteAiChat deck context', () => {
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
		const chat = new SvelteAiChat({
			bridge,
			config: { connection: { kind: 'transport', transport } },
			checkAvailable: () => Promise.resolve(true),
		});
		await chat.init();

		chat.send('hello');

		await vi.waitFor(() => expect(requests).toHaveLength(1));
		expect(JSON.stringify(requests[0])).toContain('Deck: Untitled - 2 slide(s)');
	});
});
