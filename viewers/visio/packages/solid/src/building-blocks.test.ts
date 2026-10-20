import { expect, it, vi } from 'vitest';
import { createComponent, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { DiagramCanvas, createViewerBuildingBlocks } from './building-blocks';
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
