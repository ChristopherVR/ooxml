import type { PptxAiBridge, PptxAiChatSession, PptxAiConfig } from 'ooxml-ui/pptx/ai';
import type { PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { AiPanelController } from '../../ai/ai-panel-controller.svelte';
import AiChatPanel from './AiChatPanel.svelte';

/** Runs on each session the panel builds, so a test can stage a proposal the way a tool would. */
const built = vi.hoisted(() => ({
	onSession: undefined as ((session: PptxAiChatSession) => void) | undefined,
}));

vi.mock(import('ooxml-ui/pptx/ai'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		createAiChatSession: async (...args: Parameters<typeof actual.createAiChatSession>) => {
			const session = await actual.createAiChatSession(...args);
			built.onSession?.(session);
			return session;
		},
	};
});

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	built.onSession = undefined;
});

/** A controller with no selection (live focus = the active slide). */
function fakeController(): AiPanelController {
	return new AiPanelController({
		getActiveSlideIndex: () => 0,
		getSelectedElementId: () => null,
		getSelectedElementIds: () => [],
		getSelectedElement: () => undefined,
		openPanel: () => undefined,
	});
}

/** A minimal bridge whose deck load the test controls. */
function fakeBridge(getDeckLoad: () => number = () => 1): PptxAiBridge {
	const slides = [{ id: 's1', slideNumber: 1, elements: [] }] as unknown as PptxSlide[];
	return {
		getDeckMeta: () => ({ slideCount: 1, activeSlideIndex: 0, width: 960, height: 540 }),
		getSlides: () => slides,
		getActiveSlideIndex: () => 0,
		getTheme: () => undefined,
		getHandler: () => undefined,
		getDeckLoad,
		goToSlide: () => undefined,
		selectElements: () => undefined,
		applySlidesUpdate: () => undefined,
		updateElement: () => undefined,
		applyTheme: () => undefined,
	};
}

// A `kind: 'transport'` connection with an inert transport keeps session
// bootstrap fully offline (no model, no network) for the shell assertions.
const config: PptxAiConfig = {
	connection: { kind: 'transport', transport: {} as never },
};

function mountPanel(bridge: PptxAiBridge, onclose = vi.fn()): HTMLElement {
	const target = document.createElement('div');
	const instance = mount(AiChatPanel, {
		target,
		props: { bridge, config, aiPanel: fakeController(), onclose },
	});
	cleanup = () => unmount(instance);
	return target;
}

describe('aiChatPanel', () => {
	it('renders the panel shell (assistant title + close control) on open', () => {
		const target = mountPanel(fakeBridge());

		expect(target.querySelector('[data-pptx-ai-panel]')).not.toBeNull();
		expect(target.querySelector('.pptx-svelte-ai-title')?.textContent).toBe('AI Assistant');
		expect(target.querySelector('[aria-label="Close AI assistant"]')).not.toBeNull();
	});

	it('invokes onclose when the close button is clicked', () => {
		const onclose = vi.fn();
		const target = mountPanel(fakeBridge(), onclose);

		target.querySelector<HTMLButtonElement>('[aria-label="Close AI assistant"]')?.click();
		expect(onclose).toHaveBeenCalledOnce();
	});

	it('drops a staged change on a presentation load and says why', async () => {
		let load = $state(1);
		built.onSession = (session) => {
			session.proposals.stage('Rename slide', (slides) => [
				{ ...slides[0], name: 'Renamed' } as PptxSlide,
			]);
		};
		const target = mountPanel(fakeBridge(() => load));
		await vi.waitFor(() => {
			flushSync();
			expect(target.querySelectorAll('.pptx-svelte-ai-proposals')).toHaveLength(1);
		});
		expect(target.querySelector('[data-testid="pptx-ai-proposals-discarded"]')).toBeNull();

		load += 1;
		flushSync();

		expect(target.querySelector('.pptx-svelte-ai-proposals')).toBeNull();
		const notice = target.querySelector('[data-testid="pptx-ai-proposals-discarded"]');
		expect(notice?.getAttribute('role')).toBe('status');
		expect(notice?.textContent?.trim()).toBe(
			'Suggested changes were discarded because another presentation was loaded or the slides were replaced.',
		);
	});
});
