// @vitest-own-worker: mounts a full viewer.
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, expect, test } from 'vitest';
import { defineComponent, h } from 'vue';

import { buildingBlocksContractDeck } from '../../../shared/test-utils/building-blocks-contract.mjs';
import BuildingBlocksCanvas from './BuildingBlocksCanvas.vue';
import type { ViewerBuildingBlocks } from './types';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';

let wrapper: VueWrapper | undefined;
afterEach(() => {
	wrapper?.unmount();
	wrapper = undefined;
	localStorage.clear();
	document.body.replaceChildren();
});

async function mountCanvas(canEdit = true) {
	const content = await buildingBlocksContractDeck(PptxHandler);
	let blocks: ViewerBuildingBlocks | undefined;
	const Host = defineComponent({
		setup() {
			blocks = useViewerBuildingBlocks({ content, canEdit, googleFonts: false, autosave: false });
			return () => h(BuildingBlocksCanvas, { blocks: blocks! });
		},
	});
	wrapper = mount(Host, { attachTo: document.body });
	await flushPromises();
	await expect.poll(() => blocks!.state.value.slideCount, { timeout: 15000 }).toBe(2);
	return blocks!;
}

test('the canvas shows no ribbon, rail or dialog, and reports not-ready before the deck', async () => {
	const blocks = await mountCanvas();
	expect(document.querySelector('[data-pptx-block="canvas"]')).not.toBeNull();
	expect(document.querySelector('[role="toolbar"]')).toBeNull();
	expect(blocks.viewer.value).not.toBeNull();
});

test('the shortcut chord and Ctrl+F reach the host instead of opening dialogs', async () => {
	const blocks = await mountCanvas();
	const root = document.querySelector<HTMLElement>('.pptx-vue-viewer')!;
	root.dispatchEvent(
		new KeyboardEvent('keydown', {
			key: 'f',
			code: 'KeyF',
			ctrlKey: true,
			bubbles: true,
			cancelable: true,
		}),
	);
	await flushPromises();
	expect(blocks.actions.value.findReplace.open).toBe(true);
	expect(document.querySelector('input[aria-label="Search text"]')).toBeNull();
	blocks.actions.value.shortcuts.setOpen(true);
	await flushPromises();
	expect(blocks.actions.value.pending).toEqual({ kind: 'shortcuts', open: true });
	expect(document.querySelector('[role="dialog"]')).toBeNull();
});

test('find matches outline their elements on the canvas', async () => {
	const blocks = await mountCanvas();
	blocks.actions.value.findReplace.setOpen(true);
	blocks.actions.value.findReplace.setQuery('alpha');
	await expect
		.poll(() => document.querySelectorAll('[data-pptx-find-match]').length, { timeout: 5000 })
		.toBeGreaterThan(0);
});

test('a layout the master does not list is refused', async () => {
	const blocks = await mountCanvas();
	const result = await blocks.actions.value.layout.apply('ppt/slideLayouts/nope.xml');
	expect(result.ok).toBe(false);
});
