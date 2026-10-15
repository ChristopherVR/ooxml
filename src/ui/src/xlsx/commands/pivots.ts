import {
	pivotGroupableDateFields,
	pivotGroupableFields,
	pivotSubtotalFields,
	pivotTableAt,
} from 'ooxml-core/xlsx';
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
			id: 'pivot.group-selection',
			label: 'Group Selection',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => {
				const selection = selected(ctx);
				const workbook = ctx.workbook();
				return (
					!!selection &&
					!!workbook &&
					((selection.pivot.dateGroups?.length ?? 0) > 0 ||
						pivotGroupableDateFields(workbook, selection.t.sheet, selection.pivot).length > 0)
				);
			},
			run: (ctx) => void ctx.dialogs.open('pivot-group-dates'),
		}),
		editing({
			id: 'pivot.ungroup',
			label: 'Ungroup',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => {
				const pivot = selected(ctx)?.pivot;
				return !!(pivot?.fieldGroups?.length || pivot?.dateGroups?.length);
			},
			run: (ctx) => {
				const selection = selected(ctx);
				if (selection)
					selection.t.session.refreshPivotTable(selection.t.sheet, selection.pivot.name, {
						fieldGroups: [],
						dateGroups: [],
					});
			},
		}),
		editing({
			id: 'pivot.field-settings',
			label: 'Field Settings',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => {
				const pivot = selected(ctx)?.pivot;
				return !!pivot && pivotSubtotalFields(pivot).length > 0;
			},
			run: (ctx) => void ctx.dialogs.open('pivot-field-settings'),
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
