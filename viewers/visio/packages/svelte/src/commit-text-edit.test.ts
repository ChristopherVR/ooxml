import { expect, it, vi } from 'vitest';
import { mount, unmount, tick } from 'svelte';
import { writable } from 'svelte/store';
import Harness from './BuildingBlocksHarness.svelte';
import type { BuildingBlockOptions, createViewerBuildingBlocks } from './building-blocks';

it('Svelte helper commits the inline text draft through the controller', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
	const component = mount(Harness, {
		target: host,
		props: {
			options: writable<BuildingBlockOptions>({ canEdit: true }),
			ready: (value: ReturnType<typeof createViewerBuildingBlocks>) => {
				blocks = value;
			},
		},
	});
	await tick();
	const spy = vi.spyOn(blocks.controller, 'commitTextEdit');
	expect(await blocks.commitTextEdit()).toMatchObject({ status: 'unchanged', reason: 'no-draft' });
	expect(spy).toHaveBeenCalledOnce();
	unmount(component);
	host.remove();
});
