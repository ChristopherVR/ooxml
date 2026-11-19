/**
 * The Angular chat service over the real shared session: what the model receives, and the
 * review states the cards show. Angular has no TestBed here (see `vitest.config.ts`), so the
 * templates are read from the component sources, as the other contract tests do.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { DestroyRef, Injector, runInInjectionContext } from '@angular/core';
import type { ChatTransport } from 'ai';
import type { PptxAiBridge, PptxAiUIMessage } from 'ooxml-ui/pptx/ai';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { AiChatService } from './ai-chat.service';

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

const readSource = (file: string) => readFileSync(path.join(import.meta.dirname, file), 'utf8');

describe('angular AI session wiring', () => {
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
		const injector = Injector.create({ providers: [] });
		const service = runInInjectionContext(injector, () => new AiChatService());

		service.init(bridge, { connection: { kind: 'transport', transport } });
		// The first import of the chat runtime and the `ai` SDK takes more than waitFor's default second.
		await vi.waitFor(() => expect(service.state()).toBe('ready'), { timeout: 15_000 });
		service.send('hello');

		await vi.waitFor(() => expect(requests).toHaveLength(1));
		expect(JSON.stringify(requests[0])).toContain('Deck: Untitled - 2 slide(s)');
		injector.get(DestroyRef);
		injector.destroy();
	});

	it('shows conflicts and approval needs on the proposal card', () => {
		const card = readSource('ai-proposal-card.component.ts');
		expect(card).toContain('pptx.ai.proposalConflict');
		expect(card).toContain('pptx.ai.proposalNeedsApproval');
	});

	it('says the suggested changes were discarded where they were listed', () => {
		const panel = readSource('ai-chat-panel.component.ts');
		const notice = panel.slice(panel.indexOf('@else if (chat.discarded())'));
		expect(notice).toMatch(
			/^@else if \(chat\.discarded\(\)\) \{\s*<p\s+role="status"\s+data-testid="pptx-ai-proposals-discarded"[^>]*>\s*\{\{ 'pptx\.ai\.proposalsDiscarded' \| translate \}\}/u,
		);
		expect(panel.indexOf('@if (chat.proposals().length > 0)')).toBeLessThan(
			panel.indexOf('@else if (chat.discarded())'),
		);
	});

	it('offers Accept all only for changes it can apply', () => {
		expect(readSource('ai-chat-panel.component.ts')).toContain(
			'canAcceptAllProposals(this.chat.proposals())',
		);
	});
});
