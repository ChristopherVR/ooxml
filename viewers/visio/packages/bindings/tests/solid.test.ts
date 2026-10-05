import { beforeEach, expect, it, vi } from 'vitest';
import { current, emit, reset } from './mock-binding.js';
import { createComponent, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { VisioViewer } from '../src/solid.jsx';
import type { ViewerCallbacks, ViewerHandle } from '../src/common.js';
beforeEach(reset);
it('Solid tracks props and callbacks, releases refs and destroys its native owner', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = vi.fn();
	const latest = vi.fn();
	const refs = vi.fn();
	const [zoom, setZoom] = createSignal(2);
	const [events, setEvents] = createSignal<ViewerCallbacks | undefined>({ 'zoom-change': first });
	let retained: ViewerHandle | undefined;
	const dispose = render(
		() =>
			createComponent(VisioViewer, {
				document: null,
				pageIndex: 0,
				get zoom() {
					return zoom();
				},
				showToolbar: false,
				get events() {
					return events();
				},
				viewerRef(handle) {
					refs(handle);
					if (handle) retained = handle;
				},
			}),
		host,
	);
	const live = current();
	expect(live.options).toMatchObject({ document: null, pageIndex: 0, zoom: 2, showToolbar: false });
	setZoom(3);
	setEvents({ 'zoom-change': latest });
	emit('zoom-change', 3);
	expect(current()).toBe(live);
	expect(live.options.zoom).toBe(3);
	expect(first).not.toHaveBeenCalled();
	expect(latest).toHaveBeenCalledWith(3);
	setEvents(undefined);
	expect(live.options.events).toEqual({});
	dispose();
	expect(refs).toHaveBeenLastCalledWith(undefined);
	expect(live.destroy).toHaveBeenCalledOnce();
	expect(() => retained!.fit()).toThrow('not mounted');
	host.remove();
});
it('Solid disposes the shared binding even if the user ref-release callback throws', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const failure = new Error('ref release failed');
	const dispose = render(
		() =>
			createComponent(VisioViewer, {
				viewerRef(handle) {
					if (!handle) throw failure;
				},
			}),
		host,
	);
	const live = current();
	expect(dispose).toThrow(failure);
	expect(live.destroy).toHaveBeenCalledOnce();
	host.remove();
});
it('Solid transfers the imperative ref even when the previous ref-release callback throws', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const failure = new Error('old ref callback failed');
	const latest = vi.fn();
	const initial = (handle: ViewerHandle | undefined) => {
		if (!handle) throw failure;
	};
	const [ref, setRef] = createSignal<(handle: ViewerHandle | undefined) => void>(initial);
	const dispose = render(
		() =>
			createComponent(VisioViewer, {
				get viewerRef() {
					return ref();
				},
			}),
		host,
	);
	const live = current();
	expect(() => setRef(() => latest)).toThrow(failure);
	expect(latest).toHaveBeenCalledOnce();
	expect(latest.mock.calls[0]?.[0].element).toBe(live.binding.element);
	dispose();
	expect(latest).toHaveBeenLastCalledWith(undefined);
	expect(live.destroy).toHaveBeenCalledOnce();
	host.remove();
});
