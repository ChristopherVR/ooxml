import { FIND_MATCH_ATTRIBUTE } from 'ooxml-ui/pptx';
// @vitest-own-worker: mounts a full viewer.
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { buildingBlocksContractDeck } from '../../../shared/test-utils/building-blocks-contract.mjs';
import { createViewerBuildingBlocks } from './create-viewer-building-blocks';
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
	it('marks the matched text of the host find bar, the current match apart', async () => {
		const { blocks, canvas } = await mounted();
		const marks = () =>
			[...canvas.querySelectorAll(`mark[${FIND_MATCH_ATTRIBUTE}]`)].map((mark) => [
				mark.getAttribute(FIND_MATCH_ATTRIBUTE),
				mark.textContent,
			]);
		blocks.actions.findReplace.setOpen(true);
		blocks.actions.findReplace.setQuery('alpha');
		await vi.waitFor(() =>
			expect(marks()).toEqual([
				['current', 'Alpha'],
				['match', 'alpha'],
			]),
		);
		blocks.actions.findReplace.next();
		await vi.waitFor(() =>
			expect(marks()).toEqual([
				['match', 'Alpha'],
				['current', 'alpha'],
			]),
		);
		blocks.actions.findReplace.setOpen(false);
		await vi.waitFor(() => expect(marks()).toEqual([]));
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

	it('locks the Trust Center settings from googleFonts and allowExternalContent live', async () => {
		const { blocks, content } = await mounted();
		const viewer = blocks.viewer!;
		const locks = () =>
			(
				viewer as unknown as {
					customization: { getResolved(): { lockedSettings: Record<string, unknown> } };
				}
			).customization.getResolved().lockedSettings;
		const webfonts = vi.spyOn(viewer as unknown as { refreshWebfonts(): void }, 'refreshWebfonts');
		blocks.update({ content, canEdit: true, googleFonts: false, allowExternalContent: false });
		expect(locks()).toMatchObject({
			'trust.allowFontDownloads': false,
			'trust.allowExternalContent': false,
		});
		expect(webfonts).toHaveBeenCalledTimes(1);
		// The same policy in a new options object changes nothing.
		blocks.update({ content, canEdit: true, googleFonts: false, allowExternalContent: false });
		expect(webfonts).toHaveBeenCalledTimes(1);
		blocks.update({ content, canEdit: true });
		expect(locks()['trust.allowFontDownloads']).toBeUndefined();
		expect(locks()['trust.allowExternalContent']).toBeUndefined();
		expect(blocks.viewer).toBe(viewer);
		blocks.destroy();
		vi.restoreAllMocks();
	}, 30000);
});
