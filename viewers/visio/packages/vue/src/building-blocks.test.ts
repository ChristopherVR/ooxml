import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import { expect, it, vi } from 'vitest';
import { createApp, defineComponent, h, nextTick, shallowRef } from 'vue';
import {
	DiagramCanvas,
	Inspector,
	Toolbar,
	PageNavigation,
	useViewerBuildingBlocks,
} from './building-blocks';
it('Vue owns live refs, permission and canvas lifecycle without application chrome', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const allowed = shallowRef(true);
	const className = shallowRef('first');
	const visible = shallowRef(true);
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	const app = createApp(
		defineComponent({
			setup() {
				blocks = useViewerBuildingBlocks(() => ({ canEdit: allowed.value }));
				return () =>
					h('section', [
						visible.value
							? h(DiagramCanvas, { ...blocks.canvasProps, class: className.value })
							: null,
						h(Toolbar, { toolbarProps: blocks.toolbarProps.value }),
						h(PageNavigation, { toolbarProps: blocks.toolbarProps.value }),
						h(Inspector, blocks.inspectorProps),
					]);
			},
		}),
	);
	app.mount(host);
	await nextTick();
	const canvas = host.querySelector('section > div')!;
	const viewport = canvas.shadowRoot!.querySelector('[part="canvas"]');
	expect(viewport).not.toBeNull();
	expect(canvas.shadowRoot!.querySelector('office-ui-ribbon')).toBeNull();
	blocks.toolbarProps.value.onZoomChange(2);
	await nextTick();
	expect(blocks.toolbarProps.value.zoom).toBe(2);
	allowed.value = false;
	className.value = 'second';
	await nextTick();
	expect(blocks.mode.value).toBe('view');
	expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).toBe(viewport);
	blocks.toolbarProps.value.onZoomChange(3);
	expect(blocks.state.value.zoom).toBe(3);
	visible.value = false;
	await nextTick();
	expect(canvas.shadowRoot!.children).toHaveLength(0);
	visible.value = true;
	await nextTick();
	expect(blocks.toolbarProps.value.zoom).toBe(3);
	const destroy = vi.spyOn(blocks.controller, 'destroy');
	app.unmount();
	expect(destroy).toHaveBeenCalledOnce();
	expect(canvas.shadowRoot!.children).toHaveLength(0);
	host.remove();
});

it('Vue canvas and inspector release old owners when their props are replaced', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = createBuildingBlockSession();
	const second = createBuildingBlockSession();
	const firstMount = vi.fn(first.mount);
	const secondMount = vi.fn(second.mount);
	const mount = shallowRef(firstMount);
	const controller = shallowRef(first.controller);
	const firstStop = vi.fn();
	const secondStop = vi.fn();
	const subscribeFirst = vi.spyOn(first.controller, 'subscribe').mockReturnValue(firstStop);
	const subscribeSecond = vi.spyOn(second.controller, 'subscribe').mockReturnValue(secondStop);
	const app = createApp({
		setup: () => () =>
			h('section', [
				h(DiagramCanvas, { mount: mount.value }),
				h(Inspector, { controller: controller.value }),
			]),
	});
	app.mount(host);
	await nextTick();
	const canvas = host.querySelector('section > div')!;
	const oldViewport = canvas.shadowRoot!.querySelector('[part="canvas"]');
	mount.value = secondMount;
	controller.value = second.controller;
	await nextTick();
	expect(host.querySelector('section > div')).toBe(canvas);
	expect(firstMount).toHaveBeenLastCalledWith(null);
	expect(secondMount).toHaveBeenLastCalledWith(canvas);
	expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).not.toBe(oldViewport);
	expect(firstStop).toHaveBeenCalled();
	expect(subscribeSecond).toHaveBeenCalled();
	app.unmount();
	expect(secondMount).toHaveBeenLastCalledWith(null);
	expect(secondStop).toHaveBeenCalled();
	subscribeFirst.mockRestore();
	subscribeSecond.mockRestore();
	first.destroy();
	second.destroy();
	host.remove();
});
