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
