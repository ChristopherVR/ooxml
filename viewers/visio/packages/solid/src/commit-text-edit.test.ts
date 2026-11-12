import { expect, it, vi } from 'vitest';
import { createComponent } from 'solid-js';
import { render } from 'solid-js/web';
import { DiagramCanvas, createViewerBuildingBlocks } from './building-blocks';

it('Solid primitive commits the inline text draft through the controller', async () => {
	const host = document.createElement('div');
	let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
	const dispose = render(() => {
		blocks = createViewerBuildingBlocks(() => ({ canEdit: true }));
		return createComponent(DiagramCanvas, blocks.canvasProps);
	}, host);
	const spy = vi.spyOn(blocks.controller, 'commitTextEdit');
	expect(await blocks.commitTextEdit()).toMatchObject({ status: 'unchanged', reason: 'no-draft' });
	expect(spy).toHaveBeenCalledOnce();
	dispose();
});
