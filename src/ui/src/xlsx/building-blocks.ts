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
import type { SurfaceActionName } from './surface-actions';
import { canEditNow, commitPending, formatCellsOperations } from './host-workflows';
import { findOperations } from './host-find';
import { insertFunctionOperations } from './host-function';
import { openWorkflow, workflowAvailable } from './host-open';
import { contextMenuModel, menuEntries } from './context-menu/host-api';
import type { ContextMenuKind } from './surface-workflows';

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
	const commit = (): boolean => commitPending(core);
	const run = async (id: string, arg?: unknown): Promise<boolean> => {
		if (core.loading || !commit()) return false;
		return core.commands.run(id, arg);
	};
	const canEdit = canEditNow(core);
	const hostAction = (name: SurfaceActionName) => ({
		enabled: workflowAvailable(core, name),
		run: async (): Promise<boolean> => openWorkflow(core, name),
	});
	const share = hostAction('onOpenShareDialog');
	const settings = hostAction('onOpenSettings');
	const formatCells = hostAction('onOpenFormatCells');
	const findReplace = hostAction('onOpenFindReplace');
	const insertFunction = hostAction('onOpenInsertFunction');
	const find = findOperations(core);
	const insertOps = insertFunctionOperations(core);
	const formatOps = formatCellsOperations(core);
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
			/** The host owns the validation alert; false means a refused edit just stays open (Retry). */
			validationAlert: Boolean(core.host.isConnected && core.surfaceActions.onOpenValidationAlert),
			/** The host replaces the native context menus (`onContextMenu`). */
			contextMenu: Boolean(core.host.isConnected && core.surfaceActions.onContextMenu),
			/** Finding needs no edit rights and no host callback. */
			find: Boolean(core.host.isConnected && core.session && !core.loading),
			replace: find.canReplace(),
			applyFormatCells: formatOps.canApply(),
			insertFunctionCall: insertOps.canInsert(),
		},
		/**
		 * The operations behind Format Cells, Find and Replace and Insert Function, usable without
		 * the matching `onOpen...` callback. Each re-reads read-only state, protection and the
		 * pending cell edit when it runs.
		 */
		workflows: { formatCells: formatOps, find, insertFunction: insertOps },
		/**
		 * The entries of a context menu as data, with an `execute(id)` that re-checks the item when
		 * it runs. `sheet` is the tab a `tab` menu is for. Entries whose command is not available in
		 * surface mode are listed with `enabled: false`.
		 */
		contextMenu: (menu: ContextMenuKind, sheet = core.activeSheet) => {
			const input = { kind: menu, sheet } as const;
			return contextMenuModel(ctx, input, () => menuEntries(ctx, input));
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
