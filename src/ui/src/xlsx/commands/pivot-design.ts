import { pivotColumnFields, pivotRowFields, type PivotStyleOptions } from 'ooxml-core/xlsx';
import type { Command } from 'ooxml-core/xlsx/ui';
import { multiRows, patchCommand, pivotEdit } from './pivot-context';
import { selectedPivot } from './pivots';

const styles: readonly [keyof Omit<PivotStyleOptions, 'name'>, string, boolean][] = [
	['showRowHeaders', 'Row Headers', true],
	['showColHeaders', 'Column Headers', true],
	['showRowStripes', 'Banded Rows', false],
	['showColStripes', 'Banded Columns', false],
];

export function pivotDesignCommands(): Command[] {
	return [
		...(['compact', 'outline', 'tabular'] as const).map((layout, index) =>
			patchCommand(
				`layout-${layout}`,
				['Compact Form', 'Outline Form', 'Tabular Form'][index]!,
				() => ({
					layout,
					subtotalPosition: layout === 'tabular' ? 'bottom' : 'top',
				}),
				{ checked: (ctx) => (selectedPivot(ctx)?.pivot.layout ?? 'tabular') === layout },
			),
		),
		...(['top', 'bottom', 'none'] as const).map((position, index) =>
			patchCommand(
				`subtotal-${position}`,
				['Show All Subtotals at Top', 'Show All Subtotals at Bottom', 'Do Not Show Subtotals'][
					index
				]!,
				(pivot) => ({
					subtotals: position !== 'none',
					fieldSubtotals:
						position === 'none'
							? [...pivotRowFields(pivot), ...pivotColumnFields(pivot)].map((field) => ({
									field,
									functions: [],
								}))
							: (pivot.fieldSubtotals ?? []).filter(({ functions }) => functions.length),
					...(position === 'none'
						? {}
						: {
								subtotalPosition: position,
								...(position === 'top' && (pivot.layout ?? 'tabular') === 'tabular'
									? { layout: 'outline' as const }
									: {}),
							}),
				}),
			),
		),
		...(
			[
				['off', false, false, 'Off for Rows and Columns'],
				['on', true, true, 'On for Rows and Columns'],
				['rows', true, false, 'On for Rows Only'],
				['columns', false, true, 'On for Columns Only'],
			] as const
		).map(([id, rowGrandTotals, colGrandTotals, label]) =>
			patchCommand(`totals-${id}`, label, () => ({ display: { rowGrandTotals, colGrandTotals } }), {
				checked: (ctx) => {
					const display = selectedPivot(ctx)?.pivot.display;
					return (
						(display?.rowGrandTotals ?? true) === rowGrandTotals &&
						(display?.colGrandTotals ?? true) === colGrandTotals
					);
				},
			}),
		),
		...([true, false] as const).map((repeatLabels) =>
			patchCommand(
				repeatLabels ? 'repeat-labels' : 'no-repeat-labels',
				repeatLabels ? 'Repeat All Item Labels' : 'Do Not Repeat Item Labels',
				() => ({ repeatLabels }),
				{
					enabled: (ctx) => {
						const pivot = selectedPivot(ctx)?.pivot;
						return (
							!!pivot &&
							pivot.layout !== 'compact' &&
							(!!pivot.secondRowField || pivot.valuesAxis === 'rows')
						);
					},
					unavailable: 'Select an outline or tabular PivotTable with multiple row fields.',
				},
			),
		),
		...([true, false] as const).map((blankRows) =>
			patchCommand(
				blankRows ? 'insert-blank-rows' : 'remove-blank-rows',
				blankRows ? 'Insert Blank Line After Each Item' : 'Remove Blank Line After Each Item',
				() => ({ display: { blankRows } }),
				{
					enabled: multiRows,
					unavailable: 'Select a PivotTable with multiple row fields and Values on columns.',
				},
			),
		),
		...styles.map(([flag, label, fallback]) =>
			patchCommand(
				flag,
				label,
				(pivot) => ({
					display: { style: { [flag]: !(pivot.display?.style?.[flag] ?? fallback) } },
				}),
				{
					checked: (ctx) => selectedPivot(ctx)?.pivot.display?.style?.[flag] ?? fallback,
				},
			),
		),
		pivotEdit({
			id: 'pivot.style',
			label: 'PivotTable Styles',
			value: (ctx) => selectedPivot(ctx)?.pivot.display?.style?.name ?? '',
			run: (ctx, name) => {
				const selected = selectedPivot(ctx);
				if (selected && typeof name === 'string')
					selected.t.session.updatePivotTable(selected.t.sheet, selected.pivot.name, {
						display: { style: { name } },
					});
			},
		}),
	];
}
