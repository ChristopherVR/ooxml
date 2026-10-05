import { beforeEach, expect, it, vi } from 'vitest';
import { current, emit, mocks, reset } from './mock-binding.js';
import { act, createElement, createRef, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { VisioViewer } from '../src/react.js';
import type { ViewerHandle } from '../src/common.js';
beforeEach(reset);
it('React StrictMode mounts safely, updates all props and latest callbacks, and invalidates retained handles', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const ref = createRef<ViewerHandle>();
	const first = vi.fn();
	const latest = vi.fn();
	await act(async () => {
		root.render(
			createElement(
				StrictMode,
				null,
				createElement(VisioViewer, {
					ref,
					document: null,
					pageIndex: 0,
					zoom: 2,
					showToolbar: false,
					events: { 'zoom-change': first },
				}),
			),
		);
	});
	const live = current();
	const retained = ref.current!;
	expect(live.options).toMatchObject({ document: null, pageIndex: 0, zoom: 2, showToolbar: false });
	expect(mocks.instances.filter((instance) => !instance.destroyed)).toHaveLength(1);
	await act(async () => {
		root.render(
			createElement(
				StrictMode,
				null,
				createElement(VisioViewer, {
					ref,
					zoom: 3,
					showToolbar: true,
					events: { 'zoom-change': latest },
				}),
			),
		);
	});
	expect(current()).toBe(live);
	emit('zoom-change', 3);
	expect(first).not.toHaveBeenCalled();
	expect(latest).toHaveBeenCalledWith(3);
	await act(async () => {
		root.render(createElement(StrictMode, null, createElement(VisioViewer, { ref })));
	});
	expect(current().options.events).toEqual({});
	await act(async () => {
		root.unmount();
	});
	expect(mocks.instances.every((instance) => instance.destroyed)).toBe(true);
	expect(ref.current).toBeNull();
	expect(host.children).toHaveLength(0);
	expect(() => retained.fit()).toThrow('not mounted');
	await expect(retained.load(new Uint8Array())).rejects.toThrow('not mounted');
	host.remove();
});
