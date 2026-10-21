import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import { expect, it, vi } from 'vitest';
import { createComponent, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { DiagramCanvas, Inspector, createViewerBuildingBlocks } from './building-blocks';
it('Solid accessors track controller commands and reactive permissions and release ownership', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const [allowed, setAllowed] = createSignal(true);
	let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
	const dispose = render(() => {
		blocks = createViewerBuildingBlocks(() => ({ canEdit: allowed() }));
		return createComponent(DiagramCanvas, blocks.canvasProps);
	}, host);
	const canvas = host.firstElementChild!;
	expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).not.toBeNull();
	blocks.toolbarProps().onZoomChange(2);
	expect(blocks.toolbarProps().zoom).toBe(2);
	setAllowed(false);
	expect(blocks.mode()).toBe('view');
	blocks.toolbarProps().onZoomChange(3);
	expect(blocks.state().zoom).toBe(3);
	const destroy = vi.spyOn(blocks.controller, 'destroy');
	dispose();
	expect(destroy).toHaveBeenCalledOnce();
	expect(canvas.shadowRoot!.children).toHaveLength(0);
	host.remove();
});

it('Solid swaps canvas and inspector owners reactively without replacing host elements', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = createBuildingBlockSession();
	const second = createBuildingBlockSession();
	const firstMount = vi.fn(first.mount);
	const secondMount = vi.fn(second.mount);
	const [mount, setMount] = createSignal(firstMount);
	const [controller, setController] = createSignal(first.controller);
	const firstStop = vi.fn();
	const secondStop = vi.fn();
	const firstSubscribe = vi.spyOn(first.controller, 'subscribe').mockReturnValue(firstStop);
	const secondSubscribe = vi.spyOn(second.controller, 'subscribe').mockReturnValue(secondStop);
	const dispose = render(() => {
		const section = document.createElement('section');
		section.append(
			createComponent(DiagramCanvas, {
				get mount() {
					return mount();
				},
			}) as Node,
			createComponent(Inspector, {
				get controller() {
					return controller();
				},
			}) as Node,
		);
		return section;
	}, host);
	const canvas = host.querySelector('section > div')!;
	const oldViewport = canvas.shadowRoot!.querySelector('[part="canvas"]');
	setMount(() => secondMount);
	setController(second.controller);
	expect(host.querySelector('section > div')).toBe(canvas);
	expect(firstMount).toHaveBeenLastCalledWith(null);
	expect(secondMount).toHaveBeenLastCalledWith(canvas);
	expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).not.toBe(oldViewport);
	expect(firstStop).toHaveBeenCalled();
	expect(secondSubscribe).toHaveBeenCalled();
	dispose();
	expect(secondMount).toHaveBeenLastCalledWith(null);
	expect(secondStop).toHaveBeenCalled();
	firstSubscribe.mockRestore();
	secondSubscribe.mockRestore();
	first.destroy();
	second.destroy();
	host.remove();
});
