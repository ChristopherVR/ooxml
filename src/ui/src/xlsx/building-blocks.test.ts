// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createWorkbook, getCell, loadXlsx } from 'ooxml-core/xlsx';
import type { EditorContext, GridController } from 'ooxml-core/xlsx/ui';
import { defineXlsxEditor, XlsxEditorElement } from './component';

let editing = false;
let valid = true;
vi.mock('./grid', async () => {
	const { editBridge } = await import('ooxml-core/xlsx/ui');
	return {
		mountGrid(ctx: EditorContext, host: HTMLElement) {
			host.textContent = 'grid';
			const bridge = editBridge(ctx);
			const commit = () => {
				if (valid) {
					editing = false;
					bridge.set({ editing: false, mode: 'ready' }, 'grid');
				}
				return valid;
			};
			const cancel = () => {
				editing = false;
				bridge.set({ editing: false, mode: 'ready' }, 'grid');
			};
			const stop = bridge.attach({
				begin(options) {
					editing = true;
					bridge.set(
						{
							editing: true,
							mode: options.mode ?? 'edit',
							text: options.text ?? '',
							caret: options.caret ?? 0,
						},
						'grid',
					);
					return true;
				},
				commit,
				cancel,
				changed() {},
			});
			const grid: GridController = {
				focus() {},
				scrollTo() {},
				invalidate() {},
				measureText: () => 0,
				beginEdit() {
					editing = true;
				},
				commitEdit: commit,
				cancelEdit: cancel,
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
afterEach(() => {
	document.body.replaceChildren();
	editing = false;
	valid = true;
});

function mount() {
	const editor = document.createElement('xlsx-editor') as XlsxEditorElement;
	editor.surfaceOnly = true;
	editor.workbook = createWorkbook({ sheets: ['First', 'Second'] });
	document.body.append(editor);
	return editor;
}
it('formats through the existing session without mounting application controls', async () => {
	const editor = mount();
	expect(
		editor.shadowRoot!.querySelector('office-ui-title-bar, office-ui-ribbon, office-ui-backstage'),
	).toBeNull();
	const props = editor.getToolbarProps();
	expect(await props.onBold()).toBe(true);
	expect(editor.getToolbarProps().bold).toBe(true);
	expect(editor.dirty).toBe(true);
	const loaded = await loadXlsx(await editor.saveBytes());
	expect(getCell(loaded.sheets[0]!, 0, 0)?.styleId).toBeGreaterThan(0);
	editor.readOnly = true;
	expect(await props.onItalic()).toBe(false);
	expect(editor.getToolbarProps().canEdit).toBe(false);
});
it('blocks formatting, navigation and saving while cell validation fails', async () => {
	const editor = mount();
	const props = editor.getToolbarProps();
	editing = true;
	valid = false;
	expect(await props.onBold()).toBe(false);
	expect(props.onSheetChange(1)).toBe(false);
	expect(props.onSelect('B2')).toBe(false);
	expect(() => editor.select('B2')).toThrow('validation');
	editor.setActiveSheet(1);
	expect(editor.activeSheet).toBe(0);
	await expect(editor.saveBytes()).rejects.toThrow('cell edit');
	expect(editor.getToolbarProps().bold).toBe(false);
	valid = true;
	expect(await props.onBold()).toBe(true);
	expect(editing).toBe(false);
	expect(props.onSheetChange(1)).toBe(true);
	expect(editor.activeSheet).toBe(1);
});

it('shares live formula input with the grid and permits correction after validation rejects a draft', () => {
	const editor = mount();
	let updates = 0;
	const stop = editor.subscribeEditorState(() => updates++);
	const props = editor.getToolbarProps();
	valid = false;
	expect(props.onFormulaChange('rejected')).toBe(false);
	expect(editor.getToolbarProps()).toMatchObject({ formula: 'rejected', editing: true });
	expect(props.onFormulaInput('corrected', 4)).toBe(true);
	expect(editor.getToolbarProps()).toMatchObject({
		formula: 'corrected',
		formulaCaret: 4,
		editing: true,
	});
	expect(updates).toBeGreaterThan(0);
	valid = true;
	expect(props.commitEdit()).toBe(true);
	expect(editor.getToolbarProps().editing).toBe(false);
	expect(props.onFormulaInput('cancel me')).toBe(true);
	props.onCancelEdit();
	expect(editor.getToolbarProps()).toMatchObject({ formula: '', editing: false });
	stop();
});

it('does not let retained callbacks edit the workbook after the canvas is removed', async () => {
	const editor = mount();
	const props = editor.getToolbarProps();
	editor.remove();
	expect(await props.onBold()).toBe(false);
	expect(props.onFormulaInput('detached input')).toBe(false);
	expect(editor.dirty).toBe(false);
});
