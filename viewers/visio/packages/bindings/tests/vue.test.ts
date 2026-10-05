import { beforeEach, expect, it, vi } from 'vitest';
import { current, emit, reset } from './mock-binding.js';
import { createApp, h, nextTick, reactive } from 'vue';
import { VisioViewer } from '../src/vue.js';
import { viewerOptions, type ViewerHandle, type ViewerProps } from '../src/common.js';
beforeEach(reset);
it('Vue watches native props, forwards all native events and clears callbacks on unmount', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = vi.fn();
	const latest = vi.fn();
	const output = vi.fn();
	const props = reactive<ViewerProps>({
		document: null,
		pageIndex: 0,
		zoom: 2,
		showToolbar: false,
		events: { 'zoom-change': first },
	});
	let exposed: ViewerHandle | undefined;
	const app = createApp({
		render: () =>
			h(VisioViewer, {
				...viewerOptions(props),
				ref: (value: unknown) => {
					if (value) exposed = value as ViewerHandle;
				},
				'onZoom-change': output,
			}),
	});
	app.mount(host);
	await nextTick();
	const live = current();
	const retained = exposed!;
	expect(live.options).toMatchObject({ document: null, pageIndex: 0, zoom: 2, showToolbar: false });
	props.zoom = 3;
	props.showToolbar = true;
	props.events = { 'zoom-change': latest };
	await nextTick();
	expect(current()).toBe(live);
	emit('zoom-change', 3);
	expect(first).not.toHaveBeenCalled();
	expect(latest).toHaveBeenCalledWith(3);
	expect(output).toHaveBeenCalledWith(3);
	delete props.events;
	await nextTick();
	emit('zoom-change', 4);
	expect(latest).toHaveBeenCalledOnce();
	app.unmount();
	expect(live.destroy).toHaveBeenCalledOnce();
	expect(host.children).toHaveLength(0);
	expect(() => retained.fit()).toThrow('not mounted');
	host.remove();
});
