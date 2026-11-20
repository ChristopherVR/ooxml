import type {
	PptxAiBridge,
	PptxAiConfig,
	VanillaChatController,
	VanillaChatSnapshot,
} from 'ooxml-ui/pptx/ai';
import { createAiChangeAnimator, ProposalStore } from 'ooxml-ui/pptx/ai';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createInitialViewerState, createStore } from '../state';
import { createAiPanel } from './ai-panel';
import { createAiFocusController } from './ai-panel-controller';

function slide(id: string): PptxSlide {
	return { id, slideNumber: 1, elements: [] } as unknown as PptxSlide;
}

/** A minimal in-memory bridge that records applied slide updates. */
function fakeBridge(): PptxAiBridge & { applied: string[] } {
	const applied: string[] = [];
	return {
		applied,
		getDeckMeta: () => ({ slideCount: 1, activeSlideIndex: 0, width: 960, height: 540 }),
		getSlides: () => [slide('s1')],
		getActiveSlideIndex: () => 0,
		getTheme: () => undefined,
		getHandler: () => undefined,
		getDeckLoad: () => 0,
		goToSlide: () => undefined,
		selectElements: () => undefined,
		applySlidesUpdate: (_updater, label) => {
			applied.push(label);
		},
		updateElement: () => undefined,
		applyTheme: () => undefined,
	};
}

/** A controller with a real ProposalStore but no chat/network activity. */
function fakeController(bridge: PptxAiBridge): VanillaChatController {
	const proposals = new ProposalStore(bridge);
	proposals.stage('Add a title', (slides) => slides);
	return {
		sendMessage: async () => undefined,
		regenerate: async () => undefined,
		stop: async () => undefined,
		clearError: () => undefined,
		setMessages: () => undefined,
		getSnapshot: () => ({ messages: [], status: 'ready' }),
		subscribe: () => () => undefined,
		proposals,
		changeAnimator: createAiChangeAnimator(),
		getLastFinish: () => undefined,
	};
}

const config: PptxAiConfig = {
	connection: { kind: 'transport', transport: {} as never },
};

let host: HTMLElement | null = null;

afterEach(() => {
	host?.remove();
	host = null;
});

describe('createAiPanel', () => {
	it('adds the composer only once the chat can send', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		const bridge = fakeBridge();
		const sent: string[] = [];
		let chatReady: (controller: VanillaChatController) => void = () => undefined;
		const panel = createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			createChat: () => new Promise((resolve) => (chatReady = resolve)),
		});

		// A form on the page before its submit handler would reload the whole page.
		expect(host.querySelector('.pptxv-ai-composer')).toBeNull();

		chatReady({
			...fakeController(bridge),
			sendMessage: async (text) => {
				sent.push(text);
			},
		});
		await panel;
		const input = host.querySelector<HTMLTextAreaElement>('.pptxv-ai-input');
		expect(input).toBeTruthy();
		(input as HTMLTextAreaElement).value = 'hello';
		host
			.querySelector('.pptxv-ai-composer')
			?.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
		expect(sent).toEqual(['hello']);
	});

	it('renders the composer + a staged proposal card', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		const bridge = fakeBridge();

		await createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			createChat: async () => fakeController(bridge),
		});

		expect(host.querySelector('.pptxv-ai-composer')).toBeTruthy();
		expect(host.querySelector('.pptxv-ai-input')).toBeTruthy();
		// The composer carries an explicit accessible name (parity with the
		// other four bindings' composers).
		expect(host.querySelector('.pptxv-ai-input')?.getAttribute('aria-label')).toBe(
			'Ask about this deck…',
		);
		const proposals = host.querySelector<HTMLElement>('.pptxv-ai-proposals');
		expect(proposals?.hidden).toBeFalsy();
		const card = host.querySelector('.pptxv-ai-proposal');
		expect(card).toBeTruthy();
		expect(card?.querySelector('.pptxv-ai-proposal-label')?.textContent).toBe('Add a title');
	});

	it('drops the cards and says why once another presentation loads', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		let deckLoad = 1;
		const bridge = { ...fakeBridge(), getDeckLoad: () => deckLoad };
		const listeners = new Set<() => void>();

		await createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			createChat: async () => fakeController(bridge),
			subscribeDeckLoad: (listener) => {
				listeners.add(listener);
				return () => listeners.delete(listener);
			},
		});
		expect(host.querySelector('.pptxv-ai-proposal')).toBeTruthy();
		expect(host.querySelector('[data-testid="pptx-ai-proposals-discarded"]')).toBeNull();

		deckLoad = 2;
		for (const listener of listeners) {
			listener();
		}

		expect(host.querySelector('.pptxv-ai-proposal')).toBeNull();
		const notice = host.querySelector('[data-testid="pptx-ai-proposals-discarded"]');
		expect(notice?.getAttribute('role')).toBe('status');
		expect(notice?.textContent).toBe(
			'Suggested changes were discarded because another presentation was loaded or the slides were replaced.',
		);
		expect(host.querySelector<HTMLElement>('.pptxv-ai-proposals')?.hidden).toBeFalsy();
	});

	it('renders a header "Chats" button that toggles the saved-chat menu', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		const bridge = fakeBridge();

		await createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			createChat: async () => fakeController(bridge),
		});

		const chats = host.querySelector<HTMLButtonElement>('.pptxv-ai-chats');
		expect(chats).toBeTruthy();
		expect(chats?.textContent).toBe('Chats');
		// The toggle sits in the header, before the close control.
		expect(chats?.parentElement?.classList.contains('pptxv-ai-header')).toBeTruthy();

		const menu = host.querySelector<HTMLElement>('.pptxv-ai-history');
		expect(menu).toBeTruthy();
		expect(menu?.hidden).toBeTruthy();
		chats?.click();
		expect(menu?.hidden).toBeFalsy();
		expect(menu?.querySelector('.pptxv-ai-history-title')?.textContent).toBe('Saved chats');
		expect(menu?.querySelector('.pptxv-ai-history-empty')?.textContent).toBe('No saved chats yet.');
		expect(menu?.querySelector('.pptxv-ai-history-hint')?.textContent).toBe(
			'Chats are saved in this browser.',
		);
		chats?.click();
		expect(menu?.hidden).toBeTruthy();
	});

	it('offers a header "Close AI assistant" control that calls onClose', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		const bridge = fakeBridge();
		const onClose = vi.fn();

		await createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			onClose,
			createChat: async () => fakeController(bridge),
		});

		const close = host.querySelector<HTMLButtonElement>('.pptxv-ai-close');
		expect(close).toBeTruthy();
		expect(close?.getAttribute('aria-label')).toBe('Close AI assistant');
		close?.click();
		expect(onClose).toHaveBeenCalledOnce();
	});

	it('applies the proposal through the controller store when Accept is clicked', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		const bridge = fakeBridge();
		const controller = fakeController(bridge);
		const applySpy = vi.spyOn(controller.proposals, 'apply');

		await createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			createChat: async () => controller,
		});

		const accept = host.querySelector<HTMLButtonElement>('.pptxv-ai-proposal-btn.is-accept');
		expect(accept).toBeTruthy();
		accept?.click();

		expect(applySpy).toHaveBeenCalledOnce();
		// Applying routes through the bridge's single write choke point.
		expect(bridge.applied).toStrictEqual(['Add a title']);
		// The card is cleared and the region hidden after the last proposal applies.
		expect(host.querySelector('.pptxv-ai-proposal')).toBeNull();
		expect(host.querySelector<HTMLElement>('.pptxv-ai-proposals')?.hidden).toBeTruthy();
	});

	it('navigates + highlights the canvas as a tool call runs (live focus)', async () => {
		host = document.createElement('div');
		document.body.appendChild(host);
		const bridge = fakeBridge();
		const snapshot: VanillaChatSnapshot = {
			status: 'ready',
			messages: [
				{
					id: 'm1',
					role: 'assistant',
					parts: [
						{
							type: 'tool-update_element',
							toolCallId: 'c1',
							state: 'output-available',
							input: { slideIndex: 4, elementId: 'el-2' },
						},
					],
				} as unknown as VanillaChatSnapshot['messages'][number],
			],
		};
		const controller: VanillaChatController = {
			sendMessage: async () => undefined,
			regenerate: async () => undefined,
			stop: async () => undefined,
			clearError: () => undefined,
			setMessages: () => undefined,
			getSnapshot: () => snapshot,
			subscribe: () => () => undefined,
			proposals: new ProposalStore(bridge),
			changeAnimator: createAiChangeAnimator(),
			getLastFinish: () => undefined,
		};
		const store = createStore({ ...createInitialViewerState(), currentSlide: 4 });
		const focus = createAiFocusController({ store, requestOpen: () => undefined });
		const goToSlide = vi.fn();

		await createAiPanel({
			host,
			doc: document,
			t: createTranslator('en'),
			bridge,
			config,
			controller: focus,
			goToSlide,
			createChat: async () => controller,
		});

		// The tool's slide/element drove a navigation + an active on-canvas ring.
		expect(goToSlide).toHaveBeenCalledWith(4);
		expect(focus.getHighlights()).toStrictEqual([
			{ slideIndex: 4, elementId: 'el-2', variant: 'active' },
		]);
	});
});
