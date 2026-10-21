import type { EditorCore } from './editor-core';
import { cellInputText, type HorizontalAlignment } from 'ooxml-core/xlsx';
import {
	activeRef,
	selectionRef,
	selectionFormatting,
	style,
	editBridge,
	FONT_NAMES,
	FONT_SIZES,
	numberFormatOptions,
} from 'ooxml-core/xlsx/ui';
import type { SurfaceActionContext, SurfaceActionName } from './surface-actions';

const LABELS = [
	'Bold',
	'Italic',
	'Underline',
	'Undo',
	'Redo',
	'Formula Bar',
	'Name Box',
	'Sheet',
	'Font',
	'Font Size',
	'Number Format',
	'Align Left',
	'Center',
	'Align Right',
];
const ALIGNMENTS = [
	{ value: 'left', label: 'Align Left' },
	{ value: 'center', label: 'Center' },
	{ value: 'right', label: 'Align Right' },
] as const satisfies readonly { value: HorizontalAlignment; label: string }[];

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
	/** Commits the cell edit in progress; a refusal is reported as `validation-error`. */
	const commit = (): boolean => {
		if (!core.host.isConnected) return false;
		const grid = ctx.grid();
		if (!grid?.isEditing() || grid.commitEdit()) return true;
		if (!core.editError)
			ctx.emit('validation-error', {
				message: ctx.t('The current cell edit failed validation.'),
				ref: activeRef(core.selection.get()),
			});
		return false;
	};
	const run = async (id: string, arg?: unknown): Promise<boolean> => {
		if (core.loading || !commit()) return false;
		return core.commands.run(id, arg);
	};
	const canEdit = Boolean(core.host.isConnected && core.session && !core.readOnly && !core.loading);
	const actionContext = (): SurfaceActionContext => {
		const current = core.selection.get();
		return {
			sheet: core.activeSheet,
			sheetName: core.workbook?.sheets[core.activeSheet]?.name ?? '',
			selection: selectionRef(current),
			activeCell: activeRef(current),
		};
	};
	const hostAction = (name: SurfaceActionName, needsEdit: boolean) => ({
		enabled: Boolean(core.host.isConnected && core.surfaceActions[name] && (!needsEdit || canEdit)),
		run: async (): Promise<boolean> => {
			const callback = core.surfaceActions[name];
			if (!callback || (needsEdit && !canEdit) || core.loading || !commit()) return false;
			callback(actionContext());
			return true;
		},
	});
	const share = hostAction('onOpenShareDialog', false);
	const settings = hostAction('onOpenSettings', false);
	const formatCells = hostAction('onOpenFormatCells', true);
	const findReplace = hostAction('onOpenFindReplace', false);
	const insertFunction = hostAction('onOpenInsertFunction', true);
	const canUndo = core.commands.isEnabled('edit.undo');
	const canRedo = core.commands.isEnabled('edit.redo');
	return {
		enabled: {
			bold: core.commands.isEnabled('home.bold'),
			italic: core.commands.isEnabled('home.italic'),
			underline: core.commands.isEnabled('home.underline'),
			fontFamily: core.commands.isEnabled('home.font-name'),
			fontSize: core.commands.isEnabled('home.font-size'),
			alignment: core.commands.isEnabled('home.align-left'),
			numberFormat: core.commands.isEnabled('home.number-format'),
			undo: canUndo,
			redo: canRedo,
			openShareDialog: share.enabled,
			settings: settings.enabled,
			formatCells: formatCells.enabled,
			findReplace: findReplace.enabled,
			insertFunction: insertFunction.enabled,
		},
		...selectionFormatting(core.workbook, core.activeSheet, selection.ranges),
		labels: Object.fromEntries(LABELS.map((label) => [label, ctx.t(label)])),
		/** The choices a font, size, alignment or number format control can offer. */
		options: {
			fontFamilies: [...FONT_NAMES] as string[],
			fontSizes: [...FONT_SIZES] as number[],
			alignments: ALIGNMENTS.map((a) => ({ value: a.value, label: ctx.t(a.label) })),
			numberFormats: numberFormatOptions(ctx),
		},
		canEdit,
		/** Unsaved changes since the last load, save or `markClean()`. */
		dirty: core.dirty.dirty,
		/** Sharing state: whether a room is joined, who is in it and how the connection is. */
		collaboration: core.collab.state(),
		/** The validation message of the cell edit in progress, or null. Cleared as the draft changes. */
		editError: core.editError?.message ?? null,
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
		canUndo,
		canRedo,
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
		onOpenShareDialog: share.run,
		onOpenSettings: settings.run,
		onOpenFormatCells: formatCells.run,
		onOpenFindReplace: findReplace.run,
		onOpenInsertFunction: insertFunction.run,
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
