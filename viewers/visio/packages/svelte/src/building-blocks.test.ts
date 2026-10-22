import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import SwapHarness from './CanvasSwapHarness.svelte';
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
	// Granted, but no editable drawing is open yet: the effective mode is 'view'.
	expect(blocks.controller.canEdit).toBe(true);
	expect(host.querySelector('output')!.textContent).toBe('view:2');
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

it('Svelte swaps canvas and inspector owners with store changes and releases old subscriptions', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = createBuildingBlockSession();
	const second = createBuildingBlockSession();
	const firstMount = vi.fn(first.mount);
	const secondMount = vi.fn(second.mount);
	const owner = writable({ mount: firstMount, controller: first.controller });
	const firstStop = vi.fn();
	const secondStop = vi.fn();
	const firstSubscribe = vi.spyOn(first.controller, 'subscribe').mockReturnValue(firstStop);
	const secondSubscribe = vi.spyOn(second.controller, 'subscribe').mockReturnValue(secondStop);
	const component = mount(SwapHarness, { target: host, props: { owner } });
	await tick();
	const canvas = host.querySelector('div')!;
	const oldViewport = canvas.shadowRoot!.querySelector('[part="canvas"]');
	owner.set({ mount: secondMount, controller: second.controller });
	await tick();
	expect(host.querySelector('div')).toBe(canvas);
	expect(firstMount).toHaveBeenLastCalledWith(null);
	expect(secondMount).toHaveBeenLastCalledWith(canvas);
	expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).not.toBe(oldViewport);
	expect(firstStop).toHaveBeenCalled();
	expect(secondSubscribe).toHaveBeenCalled();
	await unmount(component);
	expect(secondMount).toHaveBeenLastCalledWith(null);
	expect(secondStop).toHaveBeenCalled();
	firstSubscribe.mockRestore();
	secondSubscribe.mockRestore();
	first.destroy();
	second.destroy();
	host.remove();
});
