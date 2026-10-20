// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { act, createElement, createRef, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { useViewerBuildingBlocks, DocumentCanvas } from '../packages/bindings/src/react';
import type { EditorHandle } from '../packages/bindings/src/index';
import { createDocument, saveDocx } from '../packages/core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
it('composes a surface in StrictMode and keeps lifecycle, callbacks and permission reactive', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const handle = createRef<EditorHandle>();
	const model = createDocument();
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host(props: { canEdit: boolean }) {
		blocks = useViewerBuildingBlocks({ documentModel: model, canEdit: props.canEdit, handle });
		return createElement(DocumentCanvas, blocks.canvasProps);
	}
	try {
		await act(async () =>
			root.render(createElement(StrictMode, null, createElement(Host, { canEdit: true }))),
		);
		expect(host.querySelectorAll('docx-editor')).toHaveLength(1);
		expect(handle.current).not.toBeNull();
		expect(blocks.toolbarProps?.canEdit).toBe(true);
		await act(async () => {
			await blocks.toolbarProps!.onBold();
		});
		expect(blocks.toolbarProps?.bold).toBe(true);
		const retained = blocks.toolbarProps!;
		await act(async () =>
			root.render(createElement(StrictMode, null, createElement(Host, { canEdit: false }))),
		);
		expect(blocks.toolbarProps?.canEdit).toBe(false);
		await act(async () => {
			expect(await retained.onItalic()).toBe(false);
		});
		const bytes = await saveDocx(model);
		await act(async () => {
			await handle.current!.load(bytes);
		});
		expect(blocks.loading).toBe(false);
		expect(blocks.error).toBeNull();
		expect(blocks.mode).toBe('view');
	} finally {
		await act(async () => root.unmount());
	}
	expect(handle.current).toBeNull();
	expect(host.children).toHaveLength(0);
});
