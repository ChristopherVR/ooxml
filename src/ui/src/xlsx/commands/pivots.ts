import {
	pivotGroupableDateFields,
	pivotGroupableFields,
	pivotChartButtons,
	pivotChartTable,
	pivotSubtotalFields,
} from 'ooxml-core/xlsx';
import {
	activeChart,
	editing,
	icon,
	structureLocked,
	target,
	type Command,
} from 'ooxml-core/xlsx/ui';
import { pivotEditing, selectedPivot } from './pivot-availability';

export { selectedPivot };

const selected = selectedPivot;

export function pivotCommands(): Command[] {
	return [
		pivotEditing({
			id: 'insert.pivot-table',
			label: 'PivotTable',
			icon: icon('table'),
			lock: 'objects',
			run: (ctx) => void ctx.dialogs.open('create-pivot-table'),
		}),
		pivotEditing({
			id: 'insert.pivot-chart',
			label: 'PivotChart',
			icon: icon('chartType'),
			lock: 'objects',
			unavailable: 'Select a source range, table or PivotTable in an editable workbook.',
			enabled: (ctx) =>
				!!selected(ctx) || (!structureLocked(ctx) && target(ctx)?.ranges.length === 1),
			run: async (ctx) => (await import('../dialogs/pivot-charts')).insertPivotChart(ctx),
		}),
		pivotEditing({
			id: 'insert.recommended-pivot-tables',
			label: 'Recommended PivotTables',
			icon: icon('table'),
			lock: 'objects',
			unavailable: 'Select a source range or table in an editable workbook.',
			enabled: (ctx) =>
				!structureLocked(ctx) && (!!selected(ctx) || target(ctx)?.ranges.length === 1),
			run: async (ctx) => (await import('../dialogs/pivot-charts')).openRecommendedPivotTables(ctx),
		}),
		editing({
			id: 'chart.pivot-field-buttons',
			label: 'Field Buttons',
			icon: icon('filter'),
			lock: 'objects',
			unavailable: 'Select an editable PivotChart linked to a worksheet PivotTable.',
			enabled: (ctx) => {
				const chart = activeChart(ctx)?.chart;
				const workbook = ctx.workbook();
				return !!chart && !!workbook && !chart.editBlocked && !!pivotChartTable(workbook, chart);
			},
			checked: (ctx) => {
				const chart = activeChart(ctx)?.chart;
				return !!chart?.chartSpace?.pivotSource && pivotChartButtons(chart).size > 0;
			},
			run: (ctx) => {
				const t = target(ctx);
				const active = activeChart(ctx);
				if (t && active)
					t.session.showPivotChartButtons(
						t.sheet,
						active.index,
						pivotChartButtons(active.chart).size === 0,
					);
			},
		}),
		pivotEditing({
			id: 'pivot.filter-field',
			label: 'Filter PivotTable Field',
			icon: icon('filter'),
			lock: 'objects',
			unavailable: 'Select a PivotTable to filter its fields.',
			enabled: (ctx) => !!selected(ctx),
			run: (ctx) => void ctx.dialogs.open('pivot-field-filter'),
		}),
		pivotEditing({
			id: 'pivot.settings',
			label: 'PivotTable Settings',
			icon: icon('table'),
			lock: 'objects',
			enabled: (ctx) => !!selected(ctx),
			run: (ctx) => void ctx.dialogs.open('pivot-settings'),
		}),
		pivotEditing({
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
		pivotEditing({
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
		pivotEditing({
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
		pivotEditing({
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
					selection.t.session.updatePivotTable(selection.t.sheet, selection.pivot.name, {
						fieldGroups: [],
						dateGroups: [],
					});
			},
		}),
		pivotEditing({
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
		pivotEditing({
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
