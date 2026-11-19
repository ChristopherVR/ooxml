import type { PptxAiBridge, PptxAiConfig, VanillaChatController } from 'ooxml-ui/pptx/ai';
import { createAiChangeAnimator, ProposalStore } from 'ooxml-ui/pptx/ai';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { aiAssistants } from '../ai-assistant-registry';
import { createPptxViewer, PptxViewer } from '../PptxViewer';
import type { PptxViewerInstance } from '../types';

let active: PptxViewerInstance[] = [];

function mount(options?: ConstructorParameters<typeof PptxViewer>[1]): HTMLElement {
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, options);
	active.push(viewer);
	return container;
}

afterEach(() => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
});

// A config whose transport is only touched if the panel is actually opened, so
// the gating checks never construct a chat or hit the network.
const aiConfig: PptxAiConfig = {
	connection: { kind: 'transport', transport: {} as never },
};

describe('ai chat mounting', () => {
	it('adds no AI toggle or panel when the `ai` option is absent', () => {
		const container = mount();
		expect(container.querySelector('.pptxv-ai-toggle')).toBeNull();
		expect(container.querySelector('.pptxv-ai-panel')).toBeNull();
	});

	it('adds a title-bar toggle and a hidden panel host when `ai` is provided', () => {
		const container = mount({ ai: aiConfig });
		const toggle = container.querySelector<HTMLButtonElement>('.pptxv-ai-toggle');
		expect(toggle).toBeTruthy();
		expect(toggle?.getAttribute('aria-expanded')).toBe('false');
		const panel = container.querySelector<HTMLElement>('.pptxv-ai-panel');
		expect(panel).toBeTruthy();
		// The pane carries the neutral cross-binding marker the e2e support
		// layer (and hosts) address the assistant by.
		expect(panel?.hasAttribute('data-pptx-ai-panel')).toBeTruthy();
		// The panel is mounted but hidden until the toggle is first clicked; the
		// heavy panel builder (and the optional `ai` SDK) load lazily on open.
		expect(panel?.hidden).toBeTruthy();
		expect(panel?.childElementCount).toBe(0);
	});

	it('removes the toggle + panel on destroy', () => {
		const container = document.createElement('div');
		document.body.appendChild(container);
		const viewer = createPptxViewer(container, { ai: aiConfig });
		expect(container.querySelector('.pptxv-ai-toggle')).toBeTruthy();
		viewer.destroy();
		expect(container.querySelector('.pptxv-ai-toggle')).toBeNull();
		expect(container.querySelector('.pptxv-ai-panel')).toBeNull();
	});
});

/** The real panel over a chat that staged one proposal, with no model behind it. */
function stagedAssistant(): PptxAiAssistant {
	const chat = (bridge: PptxAiBridge): VanillaChatController => {
		const proposals = new ProposalStore(bridge);
		proposals.stage('Add a slide', (slides) => [
			...slides,
			{ ...slides[0]!, id: 'from-the-assistant', slideNumber: 2 },
		]);
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
	};
	return aiAssistants.register(aiConfig, async () => {
		const panel = await import('./ai-panel');
		return {
			...panel,
			createAiPanel: (deps) =>
				panel.createAiPanel({ ...deps, createChat: async ({ bridge }) => chat(bridge) }),
		};
	});
}

async function deckBytes(): Promise<Uint8Array> {
	const { PptxHandler } = await import('pptx-viewer-core');
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

describe('staged AI proposals across deck changes', () => {
	async function openPanel(): Promise<{ container: HTMLElement; viewer: PptxViewer }> {
		const container = mount({ ai: stagedAssistant(), editable: true });
		const viewer = active.at(-1) as PptxViewer;
		await viewer.loadFile(await deckBytes());
		container.querySelector<HTMLButtonElement>('.pptxv-ai-toggle')?.click();
		await vi.waitFor(() => {
			if (!container.querySelector('.pptxv-ai-proposal')) {
				throw new Error('no proposal card yet');
			}
		});
		return { container, viewer };
	}

	it('keeps a proposal through Insert Layout and applies it afterwards', async () => {
		const { container, viewer } = await openPanel();
		const handler = viewer.getHandler();
		viewer.toggleMasterNavigation();
		container
			.querySelector<HTMLButtonElement>('[data-testid="pptx-master-crud-addLayout"]')
			?.click();
		await vi.waitFor(() => {
			if (viewer.getHandler() === handler) {
				throw new Error('the layout is not inserted yet');
			}
		});
		viewer.toggleMasterNavigation();

		container.querySelector<HTMLButtonElement>('.pptxv-ai-proposal-btn.is-accept')?.click();
		expect(viewer.getSlideCount()).toBe(2);
		expect(container.querySelector('.pptxv-ai-proposal')).toBeNull();
	});

	it('drops the proposal and says so when another presentation loads', async () => {
		const { container, viewer } = await openPanel();
		await viewer.loadFile(await deckBytes());

		expect(container.querySelector('.pptxv-ai-proposal')).toBeNull();
		const notice = container.querySelector('[data-testid="pptx-ai-proposals-discarded"]');
		expect(notice?.getAttribute('role')).toBe('status');
	});
});
