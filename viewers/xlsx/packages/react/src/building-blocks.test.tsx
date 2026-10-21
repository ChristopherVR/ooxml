// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { act, createElement, createRef, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import {
	useViewerBuildingBlocks,
	SpreadsheetCanvas,
	FormulaBar,
	Toolbar,
	SheetTabs,
} from './building-blocks';
import type { EditorHandle } from '../../bindings/src/index';
import { mountEditor } from '../../bindings/src/index';
import { createWorkbook, saveXlsx } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
it('keeps one name box and formula input when their cell keys overlap', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const canvas = document.createElement('div');
	document.body.append(canvas);
	const binding = mountEditor(canvas, { workbook: createWorkbook() }, true);
	const editor = binding.element;
	try {
		for (const cell of ['A1', 'B20', 'C30', 'A1']) {
			editor.select(cell);
			await act(async () => root.render(createElement(FormulaBar, editor.getToolbarProps())));
			expect(host.querySelectorAll('input')).toHaveLength(2);
			expect(host.querySelectorAll('[aria-label="Name Box"]')).toHaveLength(1);
			expect(host.querySelector<HTMLInputElement>('[aria-label="Name Box"]')!.value).toBe(cell);
			expect(host.querySelectorAll('[aria-label="Formula Bar"]')).toHaveLength(1);
		}
	} finally {
		await act(async () => root.unmount());
		binding.destroy();
	}
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

async function surface() {
	const canvas = document.createElement('div');
	document.body.append(canvas);
	const binding = mountEditor(
		canvas,
		{ workbook: createWorkbook({ sheets: ['One', 'Two'] }) },
		true,
	);
	return { binding, editor: binding.element };
}

it('renders every toolbar control from the props and routes changes to the callbacks', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const { binding, editor } = await surface();
	try {
		const render = () =>
			act(async () => root.render(createElement(Toolbar, editor.getToolbarProps())));
		await render();
		const labels = [...host.querySelectorAll('button')].map((b) => b.textContent);
		expect(labels).toEqual([
			'Undo',
			'Redo',
			'Bold',
			'Italic',
			'Underline',
			'Align Left',
			'Center',
			'Align Right',
		]);
		const selects = [...host.querySelectorAll('select')].map((s) => s.getAttribute('aria-label'));
		expect(selects).toEqual(['Font', 'Font Size', 'Number Format']);
		const button = (name: string) =>
			[...host.querySelectorAll('button')].find((b) => b.textContent === name)!;
		const mousedown = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
		button('Bold').dispatchEvent(mousedown);
		// Focus (and the grid selection) stays where it was.
		expect(mousedown.defaultPrevented).toBe(true);
		expect(button('Undo').disabled).toBe(true);
		await act(async () => button('Center').click());
		await vi.waitFor(() => expect(editor.getToolbarProps().alignment).toBe('center'));
		await render();
		expect(button('Center').getAttribute('aria-pressed')).toBe('true');
		expect(button('Align Left').getAttribute('aria-pressed')).toBe('false');
		const size = host.querySelector<HTMLSelectElement>('select[aria-label="Font Size"]')!;
		await act(async () => {
			Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')!.set!.call(size, '20');
			size.dispatchEvent(new Event('change', { bubbles: true }));
		});
		await vi.waitFor(() => expect(editor.getToolbarProps().fontSize).toBe(20));
		await render();
		expect(button('Undo').disabled).toBe(false);
		editor.readOnly = true;
		await render();
		expect(button('Bold').disabled).toBe(true);
		expect(host.querySelector<HTMLSelectElement>('select[aria-label="Font"]')!.disabled).toBe(true);
		expect(button('Center').disabled).toBe(true);
	} finally {
		await act(async () => root.unmount());
		binding.destroy();
	}
});

it('shows a mixed selection as aria-pressed mixed', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const { binding, editor } = await surface();
	try {
		editor.select('A1');
		await editor.getToolbarProps().onBold();
		editor.select('A1:B1');
		await act(async () => root.render(createElement(Toolbar, editor.getToolbarProps())));
		const bold = [...host.querySelectorAll('button')].find((b) => b.textContent === 'Bold')!;
		expect(bold.getAttribute('aria-pressed')).toBe('mixed');
	} finally {
		await act(async () => root.unmount());
		binding.destroy();
	}
});

it('switches sheets from the tabs and skips hidden sheets', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const { binding, editor } = await surface();
	try {
		const render = () =>
			act(async () => root.render(createElement(SheetTabs, editor.getToolbarProps())));
		await render();
		const tabs = () => [...host.querySelectorAll<HTMLElement>('[role="tab"]')];
		expect(tabs().map((t) => t.textContent)).toEqual(['One', 'Two']);
		expect(tabs()[0]!.getAttribute('aria-selected')).toBe('true');
		await act(async () => tabs()[1]!.click());
		expect(editor.activeSheet).toBe(1);
		await render();
		expect(tabs()[1]!.getAttribute('aria-selected')).toBe('true');
		editor.workbook!.sheets[0]!.state = 'hidden';
		await render();
		expect(tabs().map((t) => t.textContent)).toEqual(['Two']);
	} finally {
		await act(async () => root.unmount());
		binding.destroy();
	}
});

it('returns dirty, selection, collaboration and editError next to the props, plus host callbacks', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const onOpenShareDialog = vi.fn();
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	const handle = createRef<EditorHandle>();
	function Host(props: { share: boolean }) {
		blocks = useViewerBuildingBlocks({
			workbook: model,
			handle,
			...(props.share ? { onOpenShareDialog } : {}),
		});
		return createElement(SpreadsheetCanvas, blocks.canvasProps);
	}
	const model = createWorkbook();
	try {
		await act(async () => root.render(createElement(Host, { share: false })));
		expect(blocks.dirty).toBe(false);
		expect(blocks.selection).toBe('A1');
		expect(blocks.collaboration.active).toBe(false);
		expect(blocks.editError).toBeNull();
		expect(blocks.toolbarProps?.enabled.share).toBe(false);
		handle.current!.share();
		expect(onOpenShareDialog).not.toHaveBeenCalled();
		await act(async () => root.render(createElement(Host, { share: true })));
		expect(blocks.toolbarProps?.enabled.share).toBe(true);
		handle.current!.share();
		expect(onOpenShareDialog).toHaveBeenCalledTimes(1);
		await act(async () => {
			await blocks.toolbarProps!.onBold();
		});
		expect(blocks.dirty).toBe(true);
		await act(async () => handle.current!.select('C3'));
		expect(blocks.selection).toBe('C3');
		await act(async () => handle.current!.markClean());
		await vi.waitFor(() => expect(blocks.dirty).toBe(false));
	} finally {
		await act(async () => root.unmount());
	}
});

it('reports a validation failure through onValidationError and the formula bar', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const onValidationError = vi.fn();
	const model = createWorkbook();
	model.sheets[0]!.dataValidations = [
		{
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '2',
			showErrorMessage: true,
			errorStyle: 'stop',
		},
	];
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host() {
		blocks = useViewerBuildingBlocks({ workbook: model, onValidationError });
		return createElement(
			'div',
			null,
			createElement(SpreadsheetCanvas, blocks.canvasProps),
			blocks.toolbarProps ? createElement(FormulaBar, blocks.toolbarProps) : null,
		);
	}
	try {
		await act(async () => root.render(createElement(Host)));
		await act(async () => {
			blocks.toolbarProps!.onFormulaInput('9');
		});
		await act(async () => {
			expect(blocks.toolbarProps!.onFormulaChange('9')).toBe(false);
		});
		expect(onValidationError).toHaveBeenCalledTimes(1);
		expect(onValidationError.mock.calls[0]![0]).toMatchObject({ ref: 'A1' });
		expect(blocks.editError).toBeTruthy();
		const bar = host.querySelector('[aria-label="Formula Bar"]')!;
		expect(bar.getAttribute('aria-invalid')).toBe('true');
		expect(bar.parentElement!.querySelector('span[role="alert"]')?.textContent).toBe(
			blocks.editError,
		);
		await act(async () => blocks.toolbarProps!.onCancelEdit());
		expect(blocks.editError).toBeNull();
	} finally {
		await act(async () => root.unmount());
	}
});
