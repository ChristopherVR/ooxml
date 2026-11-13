// @vitest-own-worker: mounts a full viewer.
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { buildingBlocksContractDeck } from '../../../shared/test-utils/building-blocks-contract.mjs';
import { createViewerBuildingBlocks } from './create-viewer-building-blocks';
import { FIND_MATCH_ATTRIBUTE } from './find-highlight';
import type { ViewerBuildingBlocksOptions } from './types';

async function mounted(options: ViewerBuildingBlocksOptions = {}) {
	const content = await buildingBlocksContractDeck(PptxHandler);
	const blocks = createViewerBuildingBlocks({ content, canEdit: true, ...options });
	const canvas = document.createElement('div');
	document.body.append(canvas);
	blocks.mountCanvas(canvas);
	await vi.waitFor(() => expect(blocks.getState().slideCount).toBe(2));
	return { blocks, canvas, content };
}

afterEach(() => {
	document.body.replaceChildren();
});

describe('vanilla building blocks, live behaviour', () => {
	it('outlines the elements holding the host find bar matches', async () => {
		const { blocks, canvas } = await mounted();
		blocks.actions.findReplace.setOpen(true);
		blocks.actions.findReplace.setQuery('alpha');
		await vi.waitFor(() =>
			expect(canvas.querySelector(`[${FIND_MATCH_ATTRIBUTE}="current"]`)).not.toBeNull(),
		);
		blocks.actions.findReplace.setOpen(false);
		await vi.waitFor(() => expect(canvas.querySelector(`[${FIND_MATCH_ATTRIBUTE}]`)).toBeNull());
		blocks.destroy();
	}, 30000);

	it('reports unknown-layout when the handler rejects a listed layout', async () => {
		const { blocks } = await mounted();
		const layout = blocks.actions.layout.available[0]!;
		expect(layout).toBeDefined();
		vi.spyOn(blocks.viewer!.getHandler()!, 'applyLayoutToSlide').mockRejectedValueOnce(
			new Error('broken layout'),
		);
		expect(await blocks.actions.layout.apply(layout.path)).toEqual({
			ok: false,
			code: 'unknown-layout',
		});
		blocks.destroy();
		vi.restoreAllMocks();
	}, 30000);

	it('applies new callbacks and permission without remounting the canvas', async () => {
		const first = vi.fn();
		const { blocks, content } = await mounted({ onDirtyChange: first });
		const viewer = blocks.viewer;
		const second = vi.fn();
		blocks.update({ content, canEdit: true, onDirtyChange: second });
		expect(blocks.actions.smartArt.insert('basicBlockList', ['A'])).toEqual({ ok: true });
		await vi.waitFor(() => expect(second).toHaveBeenCalledWith(true));
		expect(first).not.toHaveBeenCalled();
		blocks.update({ content, canEdit: false, onDirtyChange: second });
		expect(blocks.getState().mode).toBe('preview');
		expect(blocks.viewer).toBe(viewer);
		blocks.destroy();
	}, 30000);

	it('turns the shortcut chord into the host shortcut help', async () => {
		const onRequest = vi.fn();
		const { blocks } = await mounted({ onRequest });
		const internals = blocks.viewer as unknown as {
			requestHostDialog(request: { kind: 'shortcuts' }): boolean;
		};
		expect(internals.requestHostDialog({ kind: 'shortcuts' })).toBe(true);
		expect(blocks.actions.shortcuts.open).toBe(true);
		expect(onRequest).toHaveBeenCalledWith({ kind: 'shortcuts', open: true });
		blocks.destroy();
	}, 30000);
});
