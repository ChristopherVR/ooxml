// The pieces the surface building blocks share: the edit guard every host operation runs behind,
// the workflow context, and the Format Cells operations. Find and Insert Function are in
// `host-find.ts` and `host-function.ts`.
import { sheetLocked, target, activeRef, selectionRef, formatCellsState } from 'ooxml-core/xlsx/ui';
import type { EditorCore } from './editor-core';
import type { SurfaceActionContext } from './surface-actions';
import type { FormatCellsOperations } from './surface-workflows';

/** Whether an editing operation may run now (connected, loaded, editable). Read at every call. */
export function canEditNow(core: EditorCore): boolean {
	return Boolean(core.host.isConnected && core.session && !core.readOnly && !core.loading);
}

/** Commits the cell edit in progress; a refusal is reported as `validation-error`. */
export function commitPending(core: EditorCore): boolean {
	if (!core.host.isConnected) return false;
	const grid = core.ctx.grid();
	if (!grid?.isEditing() || grid.commitEdit()) return true;
	if (!core.editError)
		core.ctx.emit('validation-error', {
			message: core.ctx.t('The current cell edit failed validation.'),
			ref: activeRef(core.selection.get()),
		});
	return false;
}

export function actionContext(core: EditorCore): SurfaceActionContext {
	const current = core.selection.get();
	return {
		sheet: core.activeSheet,
		sheetName: core.workbook?.sheets[core.activeSheet]?.name ?? '',
		selection: selectionRef(current),
		activeCell: activeRef(current),
	};
}

export function formatCellsOperations(core: EditorCore): FormatCellsOperations {
	const allowed = (): boolean => canEditNow(core) && !sheetLocked(core.ctx, 'formatCells');
	return {
		getState: () => formatCellsState(core.workbook, core.activeSheet, core.selection.get().ranges),
		canApply: allowed,
		apply(patch) {
			if (!allowed() || !commitPending(core)) return false;
			// Committing can change the selection or the sheet: read both again.
			const t = target(core.ctx);
			if (!t || !allowed()) return false;
			t.session.batch('Format Cells', () => t.session.applyStyle(t.sheet, t.ranges, patch));
			core.requestRender();
			return true;
		},
	};
}
