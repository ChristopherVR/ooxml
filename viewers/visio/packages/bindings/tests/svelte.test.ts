import { beforeEach, expect, it, vi } from 'vitest';
import { current, emit, reset } from './mock-binding.js';
import { mount, unmount, tick, flushSync } from 'svelte';
import Harness from './SvelteHarness.svelte';
beforeEach(reset);
it('Svelte action receives reactive props and callbacks and disposes on native unmount', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = vi.fn();
	const latest = vi.fn();
	const harness = mount(Harness, {
		target: host,
		props: { initialEvents: { 'zoom-change': first } },
	});
	flushSync();
	await tick();
	const live = current();
	const retained = harness.getHandle();
	expect(live.options).toMatchObject({ document: null, pageIndex: 0, zoom: 2, showToolbar: false });
	harness.update(3, { 'zoom-change': latest });
	flushSync();
	await tick();
	emit('zoom-change', 3);
	expect(current()).toBe(live);
	expect(live.options.zoom).toBe(3);
	expect(first).not.toHaveBeenCalled();
	expect(latest).toHaveBeenCalledWith(3);
	harness.update(4);
	flushSync();
	await tick();
	expect(live.options.events).toEqual({});
	await unmount(harness);
	expect(live.destroy).toHaveBeenCalledOnce();
	expect(host.children).toHaveLength(0);
	expect(() => retained.fit()).toThrow('not mounted');
	host.remove();
});
