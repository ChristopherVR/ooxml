/**
 * Staged AI proposals belong to the presentation load they were made on. Slide Master
 * view commands replace the handler without loading anything, so a proposal staged
 * before Insert Layout must still be listed and apply. File > New replaces the deck in
 * place, so its proposals go and the open panel says why instead of keeping a card whose
 * Apply does nothing; a file load re-creates the panel, which still says why.
 */
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { ribbonTabButtons } from 'ooxml-ui/pptx';
import type { PptxAiBridge, PptxAiChatSession, PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';

const captured = vi.hoisted(() => ({
	bridge: null as PptxAiBridge | null,
	sessions: [] as PptxAiChatSession[],
}));

vi.mock(import('ooxml-ui/pptx/ai'), async (importOriginal) => {
	const original = await importOriginal();
	return {
		...original,
		createAiChatSession: async (bridge, config) => {
			const session = await original.createAiChatSession(bridge, config);
			// Stage once: a load re-creates the panel, and its new session must not stage again.
			if (captured.sessions.length === 0) {
				session.proposals.stage('Hide the first slide', (slides) => {
					slides[0] = { ...slides[0]!, hidden: true };
					return slides;
				});
			}
			captured.bridge = bridge;
			captured.sessions.push(session);
			return session;
		},
	};
});

async function sampleDeck(slideCount: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slideCount });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

function stubConfig(): PptxAiConfig {
	const transport = {
		sendMessages: async () => new ReadableStream(),
		reconnectToStream: async () => null,
	};
	return { connection: { kind: 'transport', transport: transport as never } };
}

async function settle(done: () => boolean): Promise<void> {
	const deadline = Date.now() + 10_000;
	while (!done() && Date.now() < deadline) {
		await flushPromises();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
	expect(done()).toBeTruthy();
}

let wrapper: VueWrapper | undefined;
afterEach(() => {
	wrapper?.unmount();
	wrapper = undefined;
	captured.bridge = null;
	captured.sessions = [];
	localStorage.clear();
});

const button = (label: string) =>
	wrapper?.findAll('button').find((candidate) => candidate.text().trim() === label);
const applyButton = () => button('Apply');
const discardedNotice = () => wrapper?.find('[data-testid="pptx-ai-proposals-discarded"]');
const setMode = (viewer: VueWrapper, mode: string) =>
	(viewer.vm as unknown as { setMode: (next: string) => void }).setMode(mode);

/** Load a deck, open the assistant and wait for the staged proposal's card. */
async function openWithProposal(): Promise<{ viewer: VueWrapper; bridge: PptxAiBridge }> {
	const viewer = mount(PowerPointViewer, {
		props: { content: await sampleDeck(2), canEdit: true, googleFonts: false, ai: stubConfig() },
		attachTo: document.body,
	});
	wrapper = viewer;
	await settle(() => (viewer.emitted('load')?.length ?? 0) === 1);
	await viewer
		.findAll('button')
		.find((candidate) => candidate.attributes('aria-label') === 'Toggle AI assistant')
		?.trigger('click');
	await settle(() => applyButton() !== undefined);
	return { viewer, bridge: captured.bridge! };
}

// Each case loads a real deck and opens the lazy assistant panel.
describe('staged AI proposals (vue)', { timeout: 20_000 }, () => {
	it('keeps a proposal through Insert Layout, and it still applies', async () => {
		const { viewer, bridge } = await openWithProposal();
		const handlerBefore = bridge.getHandler();
		const loadBefore = bridge.getDeckLoad();

		setMode(viewer, 'master');
		const insertLayout = '[data-testid="pptx-master-crud-addLayout"]';
		await settle(() => viewer.find(insertLayout).exists());
		await viewer.find(insertLayout).trigger('click');
		await settle(() => bridge.getHandler() !== handlerBefore);
		setMode(viewer, 'edit');
		await flushPromises();

		expect(bridge.getDeckLoad()).toBe(loadBefore);
		await applyButton()?.trigger('click');
		expect(bridge.getSlides()[0]?.hidden).toBeTruthy();
		expect(applyButton()).toBeUndefined();
		expect(discardedNotice()?.exists()).toBeFalsy();
	});

	it('drops the proposal on File > New and says why', async () => {
		const { viewer, bridge } = await openWithProposal();
		const loadBefore = bridge.getDeckLoad();

		ribbonTabButtons(viewer.element)
			.find((tab) => tab.textContent?.trim() === 'File')
			?.click();
		const blank = () =>
			viewer
				.findAll('button')
				.find((candidate) => candidate.text().startsWith('Blank Presentation'));
		await settle(() => blank() !== undefined);
		await blank()?.trigger('click');
		await flushPromises();

		expect(bridge.getDeckLoad()).not.toBe(loadBefore);
		expect(applyButton()).toBeUndefined();
		const notice = discardedNotice();
		expect(notice?.attributes('role')).toBe('status');
		expect(notice?.text()).toBe(
			'Suggested changes were discarded because another presentation was loaded or the slides were replaced.',
		);
	});

	it('drops the proposal when another file loads', async () => {
		const { viewer, bridge } = await openWithProposal();
		const loadBefore = bridge.getDeckLoad();
		const [first] = captured.sessions;

		await viewer.setProps({ content: await sampleDeck(3) });
		await settle(() => (viewer.emitted('load')?.length ?? 0) === 2);

		expect(bridge.getDeckLoad()).not.toBe(loadBefore);
		expect(first?.proposals.list()).toStrictEqual([]);
		expect(first?.proposals.discarded).toBeTruthy();
		await settle(() => discardedNotice()?.exists() === true);
		expect(applyButton()).toBeUndefined();
	});
});
