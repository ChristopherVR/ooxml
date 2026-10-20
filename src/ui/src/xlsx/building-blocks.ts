import type { EditorCore } from './editor-core';
import { cellInputText, type HorizontalAlignment } from 'ooxml-core/xlsx';
import {
	activeRef,
	selectionRef,
	selectionFormatting,
	style,
	editBridge,
} from 'ooxml-core/xlsx/ui';

/** Commands always re-read permission and commit cell input before changing their target. */
export function spreadsheetToolbarProps(core: EditorCore) {
	const ctx = core.ctx;
	const selection = core.selection.get();
	const bridge = editBridge(ctx);
	const draft = bridge.state();
	const onFormulaInput = (text: string, caret = text.length): boolean => {
		if (!core.host.isConnected || !core.session || core.readOnly || core.loading) return false;
		if (!bridge.state().editing && !bridge.begin({ text, caret, mode: 'edit', source: 'bar' }))
			return false;
		bridge.update(text, caret, 'bar');
		core.requestRender();
		return true;
	};
	const commit = (): boolean =>
		core.host.isConnected && (!ctx.grid()?.isEditing() || ctx.grid()!.commitEdit());
	const run = async (id: string, arg?: unknown): Promise<boolean> => {
		if (core.loading || !commit()) return false;
		return core.commands.run(id, arg);
	};
	const canEdit = Boolean(core.host.isConnected && core.session && !core.readOnly && !core.loading);
	return {
		enabled: {
			bold: core.commands.isEnabled('home.bold'),
			italic: core.commands.isEnabled('home.italic'),
			underline: core.commands.isEnabled('home.underline'),
			fontFamily: core.commands.isEnabled('home.font-name'),
			fontSize: core.commands.isEnabled('home.font-size'),
			numberFormat: core.commands.isEnabled('home.number-format'),
		},
		...selectionFormatting(core.workbook, core.activeSheet, selection.ranges),
		labels: Object.fromEntries(
			['Bold', 'Italic', 'Underline', 'Undo', 'Redo', 'Formula Bar', 'Name Box', 'Sheet'].map(
				(label) => [label, ctx.t(label)],
			),
		),
		canEdit,
		selection: selectionRef(selection),
		activeCell: activeRef(selection),
		activeSheet: core.activeSheet,
		editing: draft.editing,
		editMode: draft.mode,
		formulaCaret: draft.caret,
		sheets:
			core.workbook?.sheets.map((sheet, index) => ({
				name: sheet.name,
				index,
				state: sheet.state,
			})) ?? [],
		canUndo: core.commands.isEnabled('edit.undo'),
		canRedo: core.commands.isEnabled('edit.redo'),
		formula: draft.editing
			? draft.text
			: core.workbook
				? cellInputText(core.workbook, core.activeSheet, selection.active.row, selection.active.col)
				: '',
		onBold: () => run('home.bold'),
		onItalic: () => run('home.italic'),
		onUnderline: () => run('home.underline'),
		onFontFamilyChange: (family: string) => run('home.font-name', family),
		onFontSizeChange: (size: number) => run('home.font-size', size),
		onNumberFormatChange: (format: string) => run('home.number-format', format),
		onAlignmentChange: async (alignment: HorizontalAlignment): Promise<boolean> => {
			if (core.loading || !commit() || !core.commands.isEnabled('home.align-left')) return false;
			style(ctx, { alignment: { horizontal: alignment } });
			core.requestRender();
			return true;
		},
		onUndo: () => run('edit.undo'),
		onRedo: () => run('edit.redo'),
		onSelect: (ref: string): boolean => {
			if (!commit()) return false;
			(core.host as HTMLElement & { select(ref: string): void }).select(ref);
			return true;
		},
		onSheetChange: (index: number): boolean => {
			if (!commit()) return false;
			core.setActiveSheet(index);
			return true;
		},
		onFormulaInput,
		onFormulaChange: (text: string): boolean => onFormulaInput(text) && bridge.commit('none'),
		onCancelEdit: () => {
			bridge.cancel();
			core.requestRender();
			ctx.grid()?.focus();
		},
		commitEdit: commit,
	};
}
export type SpreadsheetToolbarProps = ReturnType<typeof spreadsheetToolbarProps>;
