// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createWorkbook, getCell } from 'ooxml-core/xlsx';
import type { EditorContext, GridController } from 'ooxml-core/xlsx/ui';
import { defineXlsxEditor, XlsxEditorElement } from './component';
import { openContextMenu } from './context-menu/index';
import { menuEntries } from './context-menu/host-api';

vi.mock('./grid', async () => {
	const { editBridge } = await import('ooxml-core/xlsx/ui');
	return {
		mountGrid(ctx: EditorContext, host: HTMLElement) {
			host.textContent = 'grid';
			const bridge = editBridge(ctx);
			let editing = false;
			const stop = bridge.attach({
				begin(options) {
					editing = true;
					bridge.set(
						{ editing: true, mode: 'edit', text: options.text ?? '', caret: options.caret ?? 0 },
						'grid',
					);
					return true;
				},
				commit() {
					editing = false;
					bridge.set({ editing: false, mode: 'ready' }, 'grid');
					return true;
				},
				cancel() {
					editing = false;
					bridge.set({ editing: false, mode: 'ready' }, 'grid');
				},
				changed() {},
			});
			const grid: GridController = {
				focus() {},
				scrollTo() {},
				invalidate() {},
				measureText: () => 0,
				beginEdit() {},
				commitEdit: () => true,
				cancelEdit() {},
				isEditing: () => editing,
				zoom: () => 100,
				setZoom() {},
			};
			ctx.attachGrid(grid);
			return stop;
		},
	};
});
defineXlsxEditor();
afterEach(() => document.body.replaceChildren());

function mount() {
	const editor = document.createElement('xlsx-editor') as XlsxEditorElement;
	editor.surfaceOnly = true;
	const workbook = createWorkbook({ sheets: ['First', 'Second'] });
	editor.workbook = workbook;
	document.body.append(editor);
	return { editor, workbook };
}

it('lists a context menu as data and re-checks an item when it runs', async () => {
	const { editor } = mount();
	const menu = editor.getToolbarProps().contextMenu('cell');
	expect(menu.items.find((i) => i.id === 'format-cells')?.enabled).toBe(false);
	expect(menu.items.find((i) => i.id === 'copy')?.label).toBe('Copy');
	// Surface mode registers few commands: the model says so instead of pretending.
	expect(menu.items.find((i) => i.id === 'insert-cells')?.enabled).toBe(false);
	expect(await menu.execute('insert-cells')).toBe(false);
	expect(await menu.execute('nope')).toBe(false);
	const tab = editor.getToolbarProps().contextMenu('tab', 1);
	expect(tab.menuSheet).toBe(1);
	expect(tab.items.find((i) => i.id === 'rename')?.enabled).toBe(false);
	editor.readOnly = true;
	expect(editor.getToolbarProps().contextMenu('row').readOnly).toBe(true);
});

it('delegates Format Cells in a menu to the host callback when installed', async () => {
	const { editor } = mount();
	const onOpenFormatCells = vi.fn();
	editor.surfaceActions = { onOpenFormatCells };
	const menu = editor.getToolbarProps().contextMenu('cell');
	expect(menu.items.find((i) => i.id === 'format-cells')?.enabled).toBe(true);
	expect(await menu.execute('format-cells')).toBe(true);
	expect(onOpenFormatCells).toHaveBeenCalledTimes(1);
	editor.readOnly = true;
	expect(await menu.execute('format-cells')).toBe(false);
	expect(onOpenFormatCells).toHaveBeenCalledTimes(1);
});

it('lets the host take over the native menu or fall through', () => {
	const { editor } = mount();
	const requests: string[] = [];
	let take = true;
	editor.surfaceActions = {
		onContextMenu: (request) => {
			requests.push(`${request.menu}:${request.x},${request.y}`);
			return take;
		},
	};
	expect(editor.getToolbarProps().enabled.contextMenu).toBe(true);
	const ctx = (editor as unknown as { core: { ctx: EditorContext } }).core.ctx;
	const input = { kind: 'cell', sheet: 0 } as const;
	const build = () => menuEntries(ctx, input);
	openContextMenu(ctx, build(), 10, 20, { offer: { input, build } });
	expect(requests).toEqual(['cell:10,20']);
	expect(ctx.root.querySelector('office-ui-context-menu')).toBeNull();
	take = false;
	openContextMenu(ctx, build(), 10, 20, { offer: { input, build } });
	expect(ctx.root.querySelector('office-ui-context-menu')).not.toBeNull();
});

it('exposes current formatting and applies a patch with the protection recheck', () => {
	const { editor, workbook } = mount();
	editor.select('A1:B1');
	const ops = editor.getToolbarProps().workflows.formatCells;
	expect(ops.getState().font.mixed).toEqual([]);
	expect(ops.apply({ font: { bold: true } })).toBe(true);
	expect(getCell(workbook.sheets[0]!, 0, 1)?.styleId ?? 0).toBeGreaterThan(0);
	expect(editor.dirty).toBe(true);
	editor.select('A1');
	ops.apply({ font: { italic: true } });
	editor.select('A1:B1');
	expect(ops.getState().font.mixed).toContain('italic');
	editor.readOnly = true;
	expect(ops.canApply()).toBe(false);
	expect(ops.apply({ font: { bold: false } })).toBe(false);
});

it('finds in a read-only workbook and replaces only when editable', () => {
	const { editor, workbook } = mount();
	const session = (editor as unknown as { core: { session: { setCellInput: Function } } }).core
		.session;
	session.setCellInput(0, 0, 0, 'apple');
	session.setCellInput(0, 2, 0, 'apple pie');
	editor.readOnly = true;
	const ops = editor.getToolbarProps().workflows.find;
	expect(editor.getToolbarProps().enabled.find).toBe(true);
	expect(editor.getToolbarProps().enabled.replace).toBe(false);
	expect(ops.findAll({ text: 'apple' })).toHaveLength(2);
	expect(ops.findNext({ text: 'apple' })).toMatchObject({ row: 2 });
	expect(ops.findNext({ text: 'apple' }, { backwards: true })).toMatchObject({ row: 0 });
	expect(ops.replaceAll({ text: 'apple' }, 'pear')).toBe(0);
	editor.readOnly = false;
	expect(ops.replaceAll({ text: 'apple' }, 'pear')).toBe(2);
	expect(getCell(workbook.sheets[0]!, 0, 0)?.value).toBe('pear');
	expect(editor.dirty).toBe(true);
});

it('does not replace locked cells on a protected sheet', () => {
	const { editor, workbook } = mount();
	const session = (editor as unknown as { core: { session: { setCellInput: Function } } }).core
		.session;
	session.setCellInput(0, 0, 0, 'x');
	workbook.sheets[0]!.protection = { sheet: true };
	const ops = editor.getToolbarProps().workflows.find;
	expect(ops.replaceAll({ text: 'x' }, 'y')).toBe(0);
	expect(getCell(workbook.sheets[0]!, 0, 0)?.value).toBe('x');
});

it('inserts a function at the caret of the formula draft and refuses plain text', () => {
	const { editor } = mount();
	const ops = editor.getToolbarProps().workflows.insertFunction;
	expect(ops.search({ text: 'sum' }).some((f) => f.name === 'SUM')).toBe(true);
	expect(ops.categories().length).toBeGreaterThan(3);
	expect(ops.insert('NOPE')).toBeUndefined();
	expect(ops.insert('SUM')).toEqual({ text: '=SUM(', caret: 5 });
	editor.getToolbarProps().onFormulaInput('=1+2', 2);
	expect(ops.insert('ABS')).toEqual({ text: '=1ABS(+2', caret: 6 });
	expect(editor.getToolbarProps().formula).toBe('=1ABS(+2');
	editor.getToolbarProps().onFormulaInput('plain', 2);
	expect(ops.insert('SUM')).toBeUndefined();
	editor.readOnly = true;
	expect(ops.canInsert()).toBe(false);
});

it('passes operations with the callback request', async () => {
	const { editor } = mount();
	const onOpenFindReplace = vi.fn();
	editor.readOnly = true;
	editor.surfaceActions = { onOpenFindReplace };
	expect(await editor.getToolbarProps().onOpenFindReplace()).toBe(true);
	const request = onOpenFindReplace.mock.calls[0]![0];
	expect(typeof request.findAll).toBe('function');
	expect(request.canReplace()).toBe(false);
});
