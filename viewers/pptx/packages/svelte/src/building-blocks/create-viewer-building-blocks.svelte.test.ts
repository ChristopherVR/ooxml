// @vitest-own-worker: mounts the real editor state.
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, tick, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	buildingBlocksContractDeck,
	runBuildingBlocksContract,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import BuildingBlocksHarness from './BuildingBlocksHarness.svelte';
import type { ViewerBuildingBlocks, ViewerBuildingBlocksOptions } from './types';

const settle = async () => {
	for (let i = 0; i < 4; i++) {
		await new Promise<void>((resolve) => {
			setTimeout(resolve, 0);
		});
		flushSync();
		await tick();
	}
};

/** Options as `$state`, so a host-style update re-runs every reader. */
class ReactiveOptions {
	value = $state.raw<ViewerBuildingBlocksOptions>({});
	constructor(initial: ViewerBuildingBlocksOptions) {
		this.value = initial;
	}
}

const cleanups: Array<() => void> = [];

/** Mount the three blocks into `target`; returns the blocks and the option setter. */
function mountBlocks(target: HTMLElement, initial: ViewerBuildingBlocksOptions) {
	const options = new ReactiveOptions(initial);
	let blocks: ViewerBuildingBlocks | undefined;
	const component = mount(BuildingBlocksHarness, {
		target,
		props: {
			getOptions: () => options.value,
			onblocks: (created: ViewerBuildingBlocks) => {
				blocks = created;
			},
		},
	});
	cleanups.push(() => void unmount(component));
	return {
		blocks: () => blocks!,
		set: (next: ViewerBuildingBlocksOptions) => {
			options.value = next;
		},
	};
}

function svelteHarness(): BuildingBlocksContractHarness {
	let mounted: ReturnType<typeof mountBlocks> | undefined;
	let host: HTMLElement | undefined;
	return {
		name: 'svelte',
		toolbar: true,
		async mount(target: HTMLElement, initial: BuildingBlocksContractOptions) {
			host = target;
			mounted = mountBlocks(target, { ...initial });
			await settle();
		},
		async update(next: BuildingBlocksContractOptions) {
			mounted!.set({ ...next });
			await settle();
		},
		actions: () => mounted!.blocks().actions,
		state: () => mounted!.blocks().state,
		async clickPictures() {
			host!.querySelector<HTMLButtonElement>('[data-control="insertImage"]')!.click();
			await settle();
		},
		settle,
		destroy: () => mounted?.blocks().destroy(),
	};
}

afterEach(() => {
	for (const cleanup of cleanups.splice(0)) {
		cleanup();
	}
	document.body.replaceChildren();
});

describe('svelte building blocks', () => {
	it('passes the shared building-blocks contract', async () => {
		await runBuildingBlocksContract({
			harness: svelteHarness(),
			expect,
			vi,
			Handler: PptxHandler,
		});
	}, 60000);

	it('mounts no ribbon, status bar or dialogs and reports not-ready before a deck loads', async () => {
		const target = document.createElement('div');
		document.body.append(target);
		const { blocks } = mountBlocks(target, { canEdit: true });
		await settle();
		expect(blocks().actions.availability.insertImage.available).toBeFalsy();
		expect(blocks().state).toMatchObject({ slideCount: 0, slideIndex: -1 });
		expect(target.querySelector('[data-pptx-block="canvas"]')).not.toBeNull();
		expect(target.querySelector('[data-pptx-block="toolbar"]')).not.toBeNull();
		expect(target.querySelector('[data-pptx-block="slides-pane"]')).not.toBeNull();
		expect(target.querySelector('[data-pptx-chrome="ribbon"]')).toBeNull();
		expect(target.querySelector('[role="dialog"]')).toBeNull();
	});

	it('turns the canvas find chord and the shortcut flag into host state', async () => {
		const content = await buildingBlocksContractDeck(PptxHandler);
		const target = document.createElement('div');
		document.body.append(target);
		const onRequest = vi.fn();
		const { blocks } = mountBlocks(target, { content, canEdit: true, onRequest });
		await vi.waitFor(() => expect(blocks().state.slideCount).toBe(2), { timeout: 20000 });
		const first = blocks().viewerState.activeSlide?.elements[0];
		blocks().viewerState.editor.select(first!.id);
		await settle();
		blocks().viewerState.controller.openHyperlink();
		await settle();
		expect(onRequest).toHaveBeenCalledWith({ kind: 'hyperlink' });
		blocks().actions.cancel();
		blocks().viewerState.parityUi.shortcutsOpen = true;
		await settle();
		expect(blocks().actions.pending).toMatchObject({ kind: 'shortcuts', open: true });
		blocks().actions.cancel();
		await settle();
		expect(blocks().viewerState.parityUi.shortcutsOpen).toBeFalsy();
		blocks().viewerState.findReplace.toggle();
		await settle();
		expect(blocks().actions.findReplace.open).toBeTruthy();
		expect(blocks().toolbar.find((group) => group.id === 'editing')?.controls[0]).toMatchObject({
			id: 'find',
			pressed: true,
		});
	}, 40000);

	it('marks the matched find text and applies hidden toolbar actions live', async () => {
		const content = await buildingBlocksContractDeck(PptxHandler);
		const target = document.createElement('div');
		document.body.append(target);
		const { blocks, set } = mountBlocks(target, { content, canEdit: true });
		await vi.waitFor(() => expect(blocks().state.slideCount).toBe(2), { timeout: 20000 });
		expect(target.querySelector('[data-control="undo"]')).not.toBeNull();
		set({ content, canEdit: true, hiddenActions: ['undo', 'redo'] });
		await settle();
		expect(target.querySelector('[data-control="undo"]')).toBeNull();
		expect(blocks().state.slideCount).toBe(2);
		blocks().actions.findReplace.setOpen(true);
		blocks().actions.findReplace.setQuery('alpha');
		const marks = () =>
			[...target.querySelectorAll('mark[data-pptx-find-match]')].map((mark) => [
				mark.getAttribute('data-pptx-find-match'),
				mark.textContent,
			]);
		await vi.waitFor(
			() =>
				expect(marks()).toEqual([
					['current', 'Alpha'],
					['match', 'alpha'],
				]),
			{ timeout: 5000 },
		);
		blocks().actions.findReplace.next();
		await vi.waitFor(() =>
			expect(marks()).toEqual([
				['match', 'Alpha'],
				['current', 'alpha'],
			]),
		);
		blocks().actions.findReplace.setOpen(false);
		await vi.waitFor(() => expect(target.querySelector('[data-pptx-find-match]')).toBeNull());
	}, 40000);
});
