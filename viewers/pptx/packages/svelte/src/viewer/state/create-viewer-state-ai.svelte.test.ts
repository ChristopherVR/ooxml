/**
 * The assistant's staged changes belong to the presentation load they were made
 * on: ordinary edits that rebuild the package (Slide Master view's Insert Layout)
 * keep them, and loading a presentation drops them. Mounts the real
 * `createViewerState`, so the bridge's deck load is the one the viewer wires.
 */
import { ProposalStore } from 'ooxml-ui/pptx/ai';
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { ViewerStateBag } from './create-viewer-state-types';
import CreateViewerStateHarness from './CreateViewerStateHarness.svelte';

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

async function buildDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

/** Wait until the editor has adopted `count` loads. */
async function waitForLoads(state: ViewerStateBag, count: number): Promise<void> {
	await vi.waitFor(
		() => {
			flushSync();
			expect(state.loader.loadCount).toBe(count);
			expect(state.editor.slides.length).toBeGreaterThan(0);
		},
		{ timeout: 30_000 },
	);
	flushSync();
}

async function loadHarness(source: Uint8Array): Promise<ViewerStateBag> {
	let captured: ViewerStateBag | undefined;
	const target = document.createElement('div');
	const instance = mount(CreateViewerStateHarness, {
		target,
		props: { source, editable: true, onready: (state: ViewerStateBag) => (captured = state) },
	});
	cleanup = () => unmount(instance);
	if (!captured) {
		throw new Error('createViewerState harness did not report its state synchronously');
	}
	await waitForLoads(captured, 1);
	return captured;
}

function stageRename(store: ProposalStore): string {
	return store.stage('Rename slide', (slides) =>
		slides.map((slide, index) => (index === 0 ? { ...slide, name: 'Renamed' } : slide)),
	).id;
}

describe('svelte AI staged changes', () => {
	it('keeps a staged change through Insert Layout, then applies it', async () => {
		const state = await loadHarness(await buildDeck());
		const store = new ProposalStore(state.ai.bridge);
		const id = stageRename(store);
		const handler = state.loader.handler;

		state.editor.masterOps.enter(0, null);
		const prompts = { promptName: () => null, notify: vi.fn(), translate: (key: string) => key };
		await expect(state.editor.masterCrud.run('addLayout', prompts)).resolves.toBeTruthy();
		state.editor.masterOps.exit();
		flushSync();

		expect(state.loader.handler).not.toBe(handler);
		expect(store.list().map((proposal) => proposal.id)).toStrictEqual([id]);
		expect(store.apply(id)).toBeTruthy();
		flushSync();
		expect(state.editor.slides[0]?.name).toBe('Renamed');
	});

	it('drops staged changes when a presentation is loaded', async () => {
		const bytes = await buildDeck();
		const state = await loadHarness(bytes);
		const store = new ProposalStore(state.ai.bridge);
		const id = stageRename(store);
		const load = state.ai.bridge.getDeckLoad();

		await state.loader.load(bytes);
		await waitForLoads(state, 2);

		expect(state.ai.bridge.getDeckLoad()).not.toBe(load);
		expect(store.list()).toStrictEqual([]);
		expect(store.discarded).toBeTruthy();
		expect(store.apply(id)).toBeFalsy();
		expect(state.editor.slides[0]?.name).not.toBe('Renamed');
	});
});
