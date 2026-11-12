import { expect, it } from 'vitest';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { DiagramCanvas, useViewerBuildingBlocks } from './building-blocks';

it('React hook commits the inline text draft through the controller', async () => {
	const host = document.createElement('div');
	const root = createRoot(host);
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host() {
		blocks = useViewerBuildingBlocks({ canEdit: true });
		return createElement(DiagramCanvas, blocks.canvasProps);
	}
	try {
		await act(async () => root.render(createElement(Host)));
		expect(await blocks.commitTextEdit()).toMatchObject({
			status: 'unchanged',
			reason: 'no-draft',
		});
		// The imperative handle is the controller and offers the same operation.
		expect(await blocks.inspectorProps.controller.commitTextEdit()).toMatchObject({
			status: 'unchanged',
		});
	} finally {
		await act(async () => root.unmount());
	}
});
