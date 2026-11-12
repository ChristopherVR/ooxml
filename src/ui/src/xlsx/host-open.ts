// Opens a host workflow: calls the host's callback with the context and the operations it needs.
import type { EditorCore } from './editor-core';
import type { SurfaceActionName } from './surface-actions';
import { findOperations } from './host-find';
import { insertFunctionOperations } from './host-function';
import { actionContext, canEditNow, commitPending, formatCellsOperations } from './host-workflows';

/** Workflows that change the workbook, so they need edit rights. Find and share do not. */
const NEEDS_EDIT: readonly SurfaceActionName[] = ['onOpenFormatCells', 'onOpenInsertFunction'];

/** Whether the workflow's host callback is installed and its edit rights are there now. */
export function workflowAvailable(core: EditorCore, name: SurfaceActionName): boolean {
	return Boolean(
		core.host.isConnected &&
		core.surfaceActions[name] &&
		(!NEEDS_EDIT.includes(name) || canEditNow(core)),
	);
}

/**
 * Calls the host callback of workflow `name` with its request. False (and nothing called) when
 * the callback is missing, the editor lacks the rights now, or a pending cell edit cannot be
 * committed.
 */
export function openWorkflow(core: EditorCore, name: SurfaceActionName): boolean {
	if (!workflowAvailable(core, name) || core.loading || !commitPending(core)) return false;
	const context = actionContext(core);
	const actions = core.surfaceActions;
	if (name === 'onOpenFormatCells') actions[name]?.({ ...context, ...formatCellsOperations(core) });
	else if (name === 'onOpenFindReplace') actions[name]?.({ ...context, ...findOperations(core) });
	else if (name === 'onOpenInsertFunction')
		actions[name]?.({ ...context, ...insertFunctionOperations(core) });
	else actions[name]?.(context);
	return true;
}
