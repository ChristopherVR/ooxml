// @vitest-own-worker: mounts a full viewer.
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, expect, test, vi } from 'vitest';
import { defineComponent, h, shallowRef } from 'vue';
import * as Y from 'yjs';

import { runBuildingBlocksContract } from '../../../shared/test-utils/building-blocks-contract.mjs';
import type { BuildingBlocksContractOptions } from '../../../shared/test-utils/building-blocks-contract.mjs';
import { runExternalContentContract } from '../../../shared/test-utils/external-content-contract.mjs';
import type { ExternalContentContractHarness } from '../../../shared/test-utils/external-content-contract.mjs';
import ExternalContentBar from '../viewer/components/ExternalContentBar.vue';
import BuildingBlocksCanvas from './BuildingBlocksCanvas.vue';
import BuildingBlocksSlidesPane from './BuildingBlocksSlidesPane.vue';
import BuildingBlocksToolbar from './BuildingBlocksToolbar.vue';
import type { ViewerBuildingBlocks } from './types';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';

let wrapper: VueWrapper | undefined;

function vueHarness(): ExternalContentContractHarness {
	const options = shallowRef<BuildingBlocksContractOptions | null>(null);
	let blocks: ViewerBuildingBlocks | undefined;
	const Host = defineComponent({
		setup() {
			blocks = useViewerBuildingBlocks(() => ({
				googleFonts: false,
				...options.value!,
				autosave: false,
			}));
			return () =>
				h('div', [
					blocks!.externalContent.value.barVisible
						? h(ExternalContentBar, {
								onEnable: blocks!.externalContent.value.enable,
								onClose: blocks!.externalContent.value.close,
							})
						: null,
					h(BuildingBlocksToolbar, { blocks: blocks! }),
					h(BuildingBlocksSlidesPane, { blocks: blocks! }),
					h(BuildingBlocksCanvas, { blocks: blocks! }),
				]);
		},
	});
	return {
		name: 'vue',
		toolbar: true,
		async mount(host, next) {
			options.value = next;
			wrapper = mount(Host, { attachTo: host });
			await flushPromises();
		},
		async update(next) {
			options.value = next;
			await flushPromises();
		},
		actions: () => blocks!.actions.value,
		state: () => blocks!.state.value,
		externalContent: () => blocks!.externalContent.value,
		slides: () => blocks!.bridge.engine.value?.slidesPane().slides ?? [],
		async clickPictures() {
			wrapper!.get('[data-control="insertImage"]').element.click();
			await flushPromises();
		},
		settle: () => flushPromises(),
		destroy: () => wrapper?.unmount(),
	};
}

afterEach(() => {
	wrapper?.unmount();
	wrapper = undefined;
	localStorage.clear();
	document.body.replaceChildren();
});

test('passes the shared security bar contract', async () => {
	await runExternalContentContract({ harness: vueHarness(), expect, Handler: PptxHandler, JSZip });
}, 60000);

test('passes the shared building-blocks contract', async () => {
	await runBuildingBlocksContract({ harness: vueHarness(), expect, vi, Handler: PptxHandler, Y });
}, 60000);
