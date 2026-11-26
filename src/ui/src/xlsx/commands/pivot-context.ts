import {
	pivotHasSlicers,
	supportsPivotCalculatedField,
	supportsPivotCondition,
	type PivotDisplayOptions,
	type PivotTable,
	type PivotTablePatch,
} from 'ooxml-core/xlsx';
import { icon, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import { selectedPivot } from './pivots';
import { pivotEditing } from './pivot-availability';

const selectedReason = 'Select an editable PivotTable.';
export const pivotConnection = (ctx: EditorContext) => {
	const selected = selectedPivot(ctx);
	return (
		!!selected && pivotHasSlicers(selected.t.workbook, selected.t.ws.sheetId, selected.pivot.name)
	);
};

export const pivotEdit = (
	command: Omit<Command, 'editing' | 'unavailable'> & { unavailable?: string },
): Command =>
	pivotEditing({
		icon: icon('table'),
		lock: 'objects',
		unavailable: selectedReason,
		...command,
		enabled: (ctx) => !!selectedPivot(ctx) && (command.enabled?.(ctx) ?? true),
	});

export function patchCommand(
	id: string,
	label: string,
	patch: (pivot: PivotTable) => Omit<PivotTablePatch, 'source'>,
	options: Pick<Command, 'checked' | 'enabled'> & { unavailable?: string } = {},
): Command {
	return pivotEdit({
		id: `pivot.${id}`,
		label,
		...options,
		run: (ctx) => {
			const selected = selectedPivot(ctx);
			if (selected)
				selected.t.session.updatePivotTable(
					selected.t.sheet,
					selected.pivot.name,
					patch(selected.pivot),
				);
		},
	});
}

function displayCommand(
	flag: keyof Omit<PivotDisplayOptions, 'style'>,
	id: string,
	label: string,
	fallback: boolean,
): Command {
	return patchCommand(
		id,
		label,
		(pivot) => ({ display: { [flag]: !(pivot.display?.[flag] ?? fallback) } }),
		{
			checked: (ctx) => selectedPivot(ctx)?.pivot.display?.[flag] ?? fallback,
		},
	);
}

export const multiRows = (ctx: EditorContext): boolean => {
	const pivot = selectedPivot(ctx)?.pivot;
	return !!pivot?.secondRowField && pivot.valuesAxis !== 'rows';
};

export function pivotAnalyzeCommands(): Command[] {
	const input = (id: 'name' | 'source' | 'move', label: string) =>
		pivotEdit({
			id: `pivot.${id}`,
			label,
			...(id !== 'move'
				? {
						enabled: (ctx: EditorContext) => !pivotConnection(ctx),
						unavailable:
							'Remove connected slicers and timelines before renaming or changing the source.',
					}
				: {}),
			run: async (ctx) => (await import('../dialogs/pivot-context')).openPivotInput(ctx, id),
		});
	return [
		input('name', 'PivotTable Name'),
		input('source', 'Change Data Source'),
		input('move', 'Move PivotTable'),
		pivotEdit({
			id: 'pivot.field-list',
			label: 'Field List',
			run: async (ctx) => {
				await (
					await import('../dialogs/pivots')
				).openPivotTable(ctx, { kind: 'settings', fieldList: true });
			},
		}),
		pivotEdit({
			id: 'pivot.refresh-all',
			label: 'Refresh All PivotTables',
			enabled: (ctx) =>
				!ctx
					.workbook()
					?.sheets.some((sheet) => sheet.pivotTables?.length && sheet.protection?.sheet),
			unavailable: 'Refresh All requires editable PivotTables on unprotected worksheets.',
			run: (ctx) => ctx.session()?.refreshAllPivotTables(),
		}),
		patchCommand(
			'clear-filters',
			'Clear Filters',
			(pivot) => ({
				condition: null,
				itemFilters: [],
				pageFilter: pivot.pageFilter ? { field: pivot.pageFilter.field } : null,
			}),
			{
				enabled: (ctx) => !pivotConnection(ctx),
				unavailable: 'Use connected slicers and timelines to clear their filters.',
			},
		),
		{
			id: 'pivot.select',
			label: 'Entire PivotTable',
			icon: icon('table'),
			unavailable: selectedReason,
			enabled: (ctx) => !!selectedPivot(ctx),
			run: (ctx) => {
				const selected = selectedPivot(ctx);
				if (selected) {
					const { location } = selected.pivot;
					ctx.selection.set({
						ranges: [location],
						active: selected.t.active,
						anchor: location.start,
					});
				}
			},
		},
		...([true, false] as const).map((collapsed) =>
			patchCommand(
				collapsed ? 'collapse-field' : 'expand-field',
				collapsed ? 'Collapse Field' : 'Expand Field',
				() => ({ display: { collapsed } }),
				{
					enabled: (ctx) =>
						multiRows(ctx) && !!selectedPivot(ctx)?.pivot.display?.collapsed !== collapsed,
					unavailable: 'Select an expandable outer row field with Values on columns.',
				},
			),
		),
		displayCommand('showDrill', 'show-drill', '+/- Buttons', true),
		displayCommand('showHeaders', 'show-headers', 'Field Headers', true),
		pivotEdit({
			id: 'pivot.condition',
			label: 'Label and Value Filters',
			enabled: (ctx) => {
				const pivot = selectedPivot(ctx)?.pivot;
				return !!pivot && supportsPivotCondition(pivot) && !pivotConnection(ctx);
			},
			unavailable:
				'Caption and value filters require an ungrouped row-only PivotTable without other filters.',
			run: async (ctx) => (await import('../dialogs/pivot-context')).openPivotCondition(ctx),
		}),
		{
			id: 'pivot.connections',
			label: 'Filter Connections',
			icon: icon('filter'),
			enabled: pivotConnection,
			unavailable: 'This PivotTable has no connected slicers or timelines.',
			run: async (ctx) => (await import('../dialogs/pivot-context')).openPivotConnections(ctx),
		},
		...(
			[
				['drill-down', 'Drill Down'],
				['drill-up', 'Drill Up'],
			] as const
		).map(([id, label]) => ({
			id: `pivot.${id}`,
			label,
			icon: icon('table'),
			enabled: () => false,
			unavailable: 'This command requires a PivotTable model the editor does not support.',
			run: () => {},
		})),
		pivotEdit({
			id: 'pivot.calculated-field',
			label: 'Calculated Field',
			enabled: (ctx) => {
				const pivot = selectedPivot(ctx)?.pivot;
				return !!pivot && supportsPivotCalculatedField(pivot);
			},
			unavailable: 'Select an ungrouped row-only PivotTable with space for another value field.',
			run: async (ctx) => (await import('../dialogs/pivot-context')).openCalculatedField(ctx),
		}),
	];
}
