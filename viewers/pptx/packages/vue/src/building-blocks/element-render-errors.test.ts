// @vitest-own-worker: mounts a full viewer.
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, expect, test, vi } from 'vitest';
import { defineComponent, h, shallowRef } from 'vue';

import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import { runElementRenderErrorContract } from '../../../shared/test-utils/element-render-error-contract.mjs';
import BuildingBlocksCanvas from './BuildingBlocksCanvas.vue';
import BuildingBlocksSlidesPane from './BuildingBlocksSlidesPane.vue';
import type { ViewerBuildingBlocks } from './types';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';

let wrapper: VueWrapper | undefined;

function vueHarness(): BuildingBlocksContractHarness {
	const options = shallowRef<BuildingBlocksContractOptions | null>(null);
	let blocks: ViewerBuildingBlocks | undefined;
	const Host = defineComponent({
		setup() {
			blocks = useViewerBuildingBlocks(() => ({ ...options.value!, autosave: false }));
			return () =>
				h('div', [
					h(BuildingBlocksSlidesPane, { blocks: blocks! }),
					h(BuildingBlocksCanvas, { blocks: blocks! }),
				]);
		},
	});
	return {
		name: 'vue',
		toolbar: false,
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
		clickPictures: () => flushPromises(),
		settle: () => flushPromises(),
		destroy: () => {
			wrapper?.unmount();
			wrapper = undefined;
		},
	};
}

afterEach(() => {
	wrapper?.unmount();
	wrapper = undefined;
	localStorage.clear();
	document.body.replaceChildren();
});

test('passes the shared element render error contract', async () => {
	await runElementRenderErrorContract({ harness: vueHarness(), expect, vi, Handler: PptxHandler });
}, 60000);
