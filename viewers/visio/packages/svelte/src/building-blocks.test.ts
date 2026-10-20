import { expect, it, vi } from 'vitest';
import { mount, unmount, tick } from 'svelte';
import { get, writable } from 'svelte/store';
import Harness from './BuildingBlocksHarness.svelte';
import type { BuildingBlockOptions, createViewerBuildingBlocks } from './building-blocks';
it('Svelte stores drive native controls, reactive options and lifecycle cleanup', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const options = writable<BuildingBlockOptions>({ canEdit: true });
	let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
	const component = mount(Harness, {
		target: host,
		props: {
			options,
			ready: (value: ReturnType<typeof createViewerBuildingBlocks>) => {
				blocks = value;
			},
		},
	});
	await tick();
	const canvas = host.querySelector('div')!;
	expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).not.toBeNull();
	get(blocks.toolbarProps).onZoomChange(2);
	await tick();
	expect(host.querySelector('output')!.textContent).toBe('edit:2');
	options.set({ canEdit: false });
	await tick();
	expect(get(blocks.mode)).toBe('view');
	get(blocks.toolbarProps).onZoomChange(3);
	await tick();
	expect(host.querySelector('output')!.textContent).toBe('view:3');
	const destroy = vi.spyOn(blocks.controller, 'destroy');
	await unmount(component);
	expect(destroy).toHaveBeenCalledOnce();
	expect(canvas.shadowRoot!.children).toHaveLength(0);
	host.remove();
});
