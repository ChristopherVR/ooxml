// @vitest-own-worker: mounts a full viewer.
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	buildingBlocksContractDeck,
	runBuildingBlocksContract,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import type { BuildingBlocksContractHarness } from '../../../shared/test-utils/building-blocks-contract.mjs';
import { createViewerBuildingBlocks } from './create-viewer-building-blocks';
import type { ViewerBuildingBlocks } from './types';

const tick = () =>
	new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});

function vanillaHarness(): BuildingBlocksContractHarness {
	let blocks: ViewerBuildingBlocks | undefined;
	let host: HTMLElement | undefined;
	return {
		name: 'vanilla',
		toolbar: true,
		async mount(target, options) {
			host = target;
			blocks = createViewerBuildingBlocks(options);
			const toolbar = document.createElement('div');
			const pane = document.createElement('div');
			const canvas = document.createElement('div');
			target.append(toolbar, pane, canvas);
			blocks.mountToolbar(toolbar);
			blocks.mountSlidesPane(pane);
			blocks.mountCanvas(canvas);
			await tick();
		},
		async update(options) {
			blocks!.update(options);
			await tick();
		},
		actions: () => blocks!.actions,
		state: () => blocks!.getState(),
		async clickPictures() {
			host!.querySelector<HTMLButtonElement>('[data-control="insertImage"]')!.click();
			await tick();
		},
		settle: tick,
		destroy: () => blocks?.destroy(),
	};
}

afterEach(() => {
	document.body.replaceChildren();
});

describe('vanilla building blocks', () => {
	it('passes the shared building-blocks contract', async () => {
		await runBuildingBlocksContract({
			harness: vanillaHarness(),
			expect,
			vi,
			Handler: PptxHandler,
		});
	}, 60000);

	it('mounts no ribbon, thumbnails or dialogs, and reports not-ready before the canvas', async () => {
		const content = await buildingBlocksContractDeck(PptxHandler);
		const blocks = createViewerBuildingBlocks({ content, canEdit: true });
		expect(blocks.viewer).toBeNull();
		expect(blocks.actions.availability.insertImage.reason).toBe('not-ready');
		expect(blocks.getState()).toMatchObject({ mode: 'preview', loading: true, slideCount: 0 });
		const canvas = document.createElement('div');
		document.body.append(canvas);
		const listener = vi.fn();
		blocks.subscribe(listener);
		const dispose = blocks.mountCanvas(canvas);
		await vi.waitFor(() => expect(blocks.getState().slideCount).toBe(2));
		expect(listener).toHaveBeenCalledWith();
		expect(canvas.querySelector('[data-pptx-viewport]')).not.toBeNull();
		expect(canvas.querySelector('.pptxv-ribbon')).toBeNull();
		expect(canvas.querySelector('.pptxv-thumbnails')).toBeNull();
		dispose();
		expect(blocks.viewer).toBeNull();
		expect(canvas.childElementCount).toBe(0);
		blocks.destroy();
	}, 30000);

	it('turns the canvas find chord into the host find bar', async () => {
		const content = await buildingBlocksContractDeck(PptxHandler);
		const blocks = createViewerBuildingBlocks({ content, canEdit: true });
		const canvas = document.createElement('div');
		document.body.append(canvas);
		blocks.mountCanvas(canvas);
		await vi.waitFor(() => expect(blocks.getState().slideCount).toBe(2));
		const internals = blocks.viewer as unknown as {
			requestHostDialog(request: { kind: 'find' }): boolean;
		};
		expect(internals.requestHostDialog({ kind: 'find' })).toBe(true);
		expect(blocks.actions.findReplace.open).toBe(true);
		blocks.destroy();
	}, 30000);
});
