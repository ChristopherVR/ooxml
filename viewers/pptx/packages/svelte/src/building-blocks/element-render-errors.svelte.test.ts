// @vitest-own-worker: mounts the real editor state.
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, tick, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import { runElementRenderErrorContract } from '../../../shared/test-utils/element-render-error-contract.mjs';
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

function svelteHarness(): BuildingBlocksContractHarness {
	let options: ReactiveOptions | undefined;
	let blocks: ViewerBuildingBlocks | undefined;
	let component: ReturnType<typeof mount> | undefined;
	return {
		name: 'svelte',
		toolbar: false,
		async mount(target: HTMLElement, initial: BuildingBlocksContractOptions) {
			options = new ReactiveOptions({ ...initial });
			const live = options;
			component = mount(BuildingBlocksHarness, {
				target,
				props: {
					getOptions: () => live.value,
					onblocks: (created: ViewerBuildingBlocks) => {
						blocks = created;
					},
				},
			});
			await settle();
		},
		async update(next: BuildingBlocksContractOptions) {
			options!.value = { ...next };
			await settle();
		},
		actions: () => blocks!.actions,
		state: () => blocks!.state,
		clickPictures: settle,
		settle,
		destroy: () => {
			if (component) {
				void unmount(component);
				component = undefined;
			}
		},
	};
}

afterEach(() => {
	document.body.replaceChildren();
});

describe('svelte building blocks element render errors', () => {
	it('passes the shared element render error contract', async () => {
		await runElementRenderErrorContract({
			harness: svelteHarness(),
			expect,
			vi,
			Handler: PptxHandler,
		});
	}, 60000);
});
