import { pivotShareable, pivotTableAt } from 'ooxml-core/xlsx';
import { editing, target, type EditorContext } from 'ooxml-core/xlsx/ui';

const tooLargeReason = 'This PivotTable is too large to share. Stop sharing to edit it.';

/** The PivotTable under the active cell, when no drawing is selected. */
export function selectedPivot(ctx: EditorContext) {
	if (ctx.selection.get().drawing !== undefined) return;
	const t = target(ctx);
	const pivot = t ? pivotTableAt(t.ws, t.active) : undefined;
	return t && pivot ? { t, pivot } : undefined;
}

/**
 * PivotTables are shared in collaboration, except one whose definition and cache are over the
 * room's size limit: it stays on this window, so its edits wait until sharing ends.
 */
function unsharedPivot(ctx: EditorContext): boolean {
	if (!ctx.collaborationActive?.()) return false;
	const pivot = selectedPivot(ctx)?.pivot;
	return !!pivot && !pivotShareable(pivot);
}

export function assertPivotEditingAvailable(ctx: EditorContext): void {
	if (unsharedPivot(ctx)) throw new Error(tooLargeReason);
}

export function pivotEditing(
	command: Omit<Parameters<typeof editing>[0], 'unavailable'> & { unavailable?: string },
) {
	return editing({
		...command,
		enabled: (ctx) => !unsharedPivot(ctx) && (command.enabled?.(ctx) ?? true),
		unavailable: (ctx) => (unsharedPivot(ctx) ? tooLargeReason : command.unavailable),
	});
}
