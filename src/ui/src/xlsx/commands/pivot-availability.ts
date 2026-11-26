import { editing, type EditorContext } from 'ooxml-core/xlsx/ui';

const sharedReason = 'PivotTable edits are unavailable during collaboration.';

export function assertPivotEditingAvailable(ctx: EditorContext): void {
	if (ctx.collaborationActive?.()) throw new Error(sharedReason);
}

export function pivotEditing(
	command: Omit<Parameters<typeof editing>[0], 'unavailable'> & { unavailable?: string },
) {
	return editing({
		...command,
		enabled: (ctx) => !ctx.collaborationActive?.() && (command.enabled?.(ctx) ?? true),
		unavailable: (ctx) => (ctx.collaborationActive?.() ? sharedReason : command.unavailable),
	});
}
