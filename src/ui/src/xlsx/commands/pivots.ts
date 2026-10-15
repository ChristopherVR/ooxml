import { pivotGroupableFields, pivotTableAt } from 'ooxml-core/xlsx';
import { editing, icon, target, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';

function selected(ctx: EditorContext) {
	const t = target(ctx);
	const pivot = t ? pivotTableAt(t.ws, t.range.start) : undefined;
	return t && pivot ? { t, pivot } : undefined;
}

export function pivotCommands(): Command[] {
	return [
		editing({
			id: 'insert.pivot-table',
			label: 'PivotTable',
			icon: icon('table'),
			lock: 'objects',
			run: (ctx) => void ctx.dialogs.open('create-pivot-table'),
		}),
		editing({
			id: 'pivot.settings',
			label: 'PivotTable Settings',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => !!selected(ctx),
			run: (ctx) => void ctx.dialogs.open('pivot-settings'),
		}),
		editing({
			id: 'pivot.refresh',
			label: 'Refresh PivotTable',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => !!selected(ctx),
			run: (ctx) => {
				const selection = selected(ctx);
				if (selection)
					selection.t.session.refreshPivotTable(selection.t.sheet, selection.pivot.name);
			},
		}),
		editing({
			id: 'pivot.group',
			label: 'Group Field',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => {
				const selection = selected(ctx);
				return !!selection && pivotGroupableFields(selection.pivot).length > 0;
			},
			run: (ctx) => void ctx.dialogs.open('pivot-group'),
		}),
		editing({
			id: 'pivot.ungroup',
			label: 'Ungroup',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => !!selected(ctx)?.pivot.fieldGroups?.length,
			run: (ctx) => {
				const selection = selected(ctx);
				if (selection)
					selection.t.session.refreshPivotTable(selection.t.sheet, selection.pivot.name, {
						fieldGroups: [],
					});
			},
		}),
		editing({
			id: 'pivot.delete',
			label: 'Delete PivotTable',
			icon: icon('clear'),
			lock: 'objects',
			enabled: (ctx) => !!selected(ctx),
			run: (ctx) => {
				const selection = selected(ctx);
				if (selection)
					selection.t.session.deletePivotTable(selection.t.sheet, selection.pivot.name);
			},
		}),
	];
}
