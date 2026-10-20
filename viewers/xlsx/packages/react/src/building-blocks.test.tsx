// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { act, createElement, createRef, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { useViewerBuildingBlocks, SpreadsheetCanvas } from './building-blocks';
import type { EditorHandle } from '../../bindings/src/index';
import { createWorkbook, saveXlsx } from '../../core/src/index';

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
	const model = createWorkbook();
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host(props: { canEdit: boolean }) {
		blocks = useViewerBuildingBlocks({ workbook: model, canEdit: props.canEdit, handle });
		return createElement(SpreadsheetCanvas, blocks.canvasProps);
	}
	try {
		await act(async () =>
			root.render(createElement(StrictMode, null, createElement(Host, { canEdit: true }))),
		);
		expect(host.querySelectorAll('xlsx-editor')).toHaveLength(1);
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
		await act(async () => {
			await expect(handle.current!.load(new Uint8Array([0x50, 0x4b, 3, 4]))).rejects.toThrow(/./);
		});
		expect(blocks.error).not.toBeNull();
		const bytes = await saveXlsx(model);
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

it('reloads bytes into a replaced canvas without creating a second engine', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const handle = createRef<EditorHandle>();
	const model = createWorkbook({ sheets: ['Remounted content'] });

	const bytes = await saveXlsx(model);
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host(props: { canvasKey: number }) {
		blocks = useViewerBuildingBlocks({ bytes, handle });
		return createElement(SpreadsheetCanvas, { ...blocks.canvasProps, key: props.canvasKey });
	}
	try {
		for (const canvasKey of [1, 2]) {
			await act(async () => root.render(createElement(Host, { canvasKey })));
			await act(async () => {
				await vi.waitFor(() =>
					expect(JSON.stringify(handle.current!.element.workbook)).toContain('Remounted content'),
				);
			});
			expect(host.querySelectorAll('xlsx-editor')).toHaveLength(1);
			expect(blocks.loading).toBe(false);
			await act(async () => {
				expect(await blocks.toolbarProps!.onBold()).toBe(true);
			});
			expect(blocks.toolbarProps?.bold).toBe(true);
		}
	} finally {
		await act(async () => root.unmount());
	}
});
