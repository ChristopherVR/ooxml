import { expect, it } from 'vitest';
import { act, createElement, createRef, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { DiagramCanvas, useViewerBuildingBlocks } from '../src/react';
import { ViewerController } from 'ooxml-ui/visio';

it('renders the standalone canvas in StrictMode and exposes the existing controller', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const handle = createRef<ViewerController>();
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host(props: { canEdit: boolean }) {
		blocks = useViewerBuildingBlocks({ canEdit: props.canEdit, handle });
		return createElement(DiagramCanvas, blocks.canvasProps);
	}
	try {
		await act(async () =>
			root.render(createElement(StrictMode, null, createElement(Host, { canEdit: false }))),
		);
		expect(handle.current).toBeInstanceOf(ViewerController);
		expect(blocks.mode).toBe('view');
		const canvas = host.firstElementChild!.shadowRoot!;
		expect(canvas.querySelector('[part="canvas"]')).not.toBeNull();
		expect(
			canvas.querySelector('office-ui-ribbon, office-ui-title-bar, office-ui-backstage'),
		).toBeNull();
		await act(async () => blocks.toolbarProps.onZoomChange(2));
		expect(blocks.toolbarProps.zoom).toBe(2);
		await act(async () =>
			root.render(createElement(StrictMode, null, createElement(Host, { canEdit: true }))),
		);
		expect(handle.current!.canEdit).toBe(true);
	} finally {
		await act(async () => root.unmount());
		host.remove();
	}
	expect(handle.current).toBeNull();
});
