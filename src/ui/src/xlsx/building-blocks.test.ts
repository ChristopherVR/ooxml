// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createWorkbook, getCell, loadXlsx } from 'ooxml-core/xlsx';
import type { EditorContext, GridController } from 'ooxml-core/xlsx/ui';
import { defineXlsxEditor, XlsxEditorElement } from './component';

let editing = false;
let valid = true;
vi.mock('./grid', () => ({
	mountGrid(ctx: EditorContext, host: HTMLElement) {
		host.textContent = 'grid';
		const grid: GridController = {
			focus() {},
			scrollTo() {},
			invalidate() {},
			measureText: () => 0,
			beginEdit() {
				editing = true;
			},
			commitEdit() {
				if (valid) editing = false;
				return valid;
			},
			cancelEdit() {
				editing = false;
			},
			isEditing: () => editing,
			zoom: () => 100,
			setZoom() {},
		};
		ctx.attachGrid(grid);
		return () => {};
	},
}));
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
