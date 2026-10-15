/**
 * Table Layout: PowerPoint's command buttons for the selected table, expressed as one-item
 * gallery modules (the SmartArt Design approach) so every binding renders and dispatches them
 * through the gallery path it already has: descriptor, `gallery-pick`, element patch + history.
 *
 * The cells the commands act on come from `RibbonGalleryContext.selectedTableCells`; with none
 * the table frame itself is selected and alignment applies to every cell. Entries the engine or
 * the ribbon cannot back are present but disabled and say why in their tooltip, never inert.
 *
 * @module render/ribbon-galleries/table-layout-commands
 */
import type { PptxTableData, TablePptxElement } from 'ooxml-core/pptx';

import type { RibbonGalleryModule } from './gallery-module';
import type {
	RibbonGalleryApplyResult,
	RibbonGalleryContext,
	RibbonGalleryId,
} from './gallery-types';
import type { TableLayoutCell } from './table-layout-ops';
import {
	canMergeSelection,
	canSplitSelection,
	deleteSelected,
	distribute,
	insertNextToSelection,
	mergeSelection,
	splitSelection,
	styleCells,
} from './table-layout-ops';

interface Targets {
	element: TablePptxElement;
	data: PptxTableData;
	cells: readonly TableLayoutCell[];
}

interface CommandSpec {
	id: RibbonGalleryId;
	key: string;
	label: string;
	iconPath: string;
	large?: boolean;
	/** Why the command is unavailable no matter the selection. */
	unavailable?: { key: string; text: string };
	/** Extra gate on the selection (merge needs a rectangle, delete needs cells). */
	enabled?: (t: Targets) => boolean;
	run?: (t: Targets) => RibbonGalleryApplyResult | null;
}

const NEEDS_CELLS = {
	key: 'pptx.gallery.tableLayout.needsCells',
	text: 'Select cells in the table to use this command',
};
const NO_ENGINE = {
	key: 'pptx.gallery.smartArtCommand.unsupported',
	text: 'Not available yet in this viewer',
};
const IN_INSPECTOR = {
	key: 'pptx.gallery.tableLayout.inInspector',
	text: 'Set this in the Inspector while the table is selected',
};
const IN_HOME = {
	key: 'pptx.gallery.tableLayout.inHome',
	text: 'Use Home > Arrange',
};

const hasCells = (t: Targets): boolean => t.cells.length > 0;

const align = (key: 'left' | 'center' | 'right'): CommandSpec['run'] => {
	return (t) => styleCells(t.element, t.cells, { align: key });
};
const valign = (key: 'top' | 'middle' | 'bottom'): CommandSpec['run'] => {
	return (t) => styleCells(t.element, t.cells, { vAlign: key });
};

const SPECS: readonly CommandSpec[] = [
	{
		id: 'tableSelect',
		key: 'select',
		label: 'Select',
		iconPath: 'M3 3h14v14H3zM3 8h14M8 3v14',
		large: true,
		unavailable: {
			key: 'pptx.gallery.tableLayout.selectHint',
			text: 'Click or drag cells in the table to select them',
		},
	},
	{
		id: 'tableViewGridlines',
		key: 'viewGridlines',
		label: 'View Gridlines',
		iconPath: 'M3 3h14v14H3zM3 10h14M10 3v14',
		unavailable: NO_ENGINE,
	},
	{
		id: 'tableDeleteRows',
		key: 'deleteRows',
		label: 'Delete Rows',
		iconPath: 'M3 5h14v4H3zM3 13h14M7 15l6 3M13 15l-6 3',
		enabled: hasCells,
		run: (t) => deleteSelected(t.element, t.cells, 'row'),
	},
	{
		id: 'tableDeleteColumns',
		key: 'deleteColumns',
		label: 'Delete Columns',
		iconPath: 'M5 3h4v14H5zM13 3v14M15 7l3 6M18 7l-3 6',
		enabled: hasCells,
		run: (t) => deleteSelected(t.element, t.cells, 'column'),
	},
	{
		id: 'tableInsertAbove',
		key: 'insertAbove',
		label: 'Insert Above',
		iconPath: 'M3 9h14v8H3zM3 13h14M10 2v5M7.5 4.5h5',
		run: (t) => insertNextToSelection(t.element, t.cells, 'above'),
	},
	{
		id: 'tableInsertBelow',
		key: 'insertBelow',
		label: 'Insert Below',
		iconPath: 'M3 3h14v8H3zM3 7h14M10 13v5M7.5 15.5h5',
		run: (t) => insertNextToSelection(t.element, t.cells, 'below'),
	},
	{
		id: 'tableInsertLeft',
		key: 'insertLeft',
		label: 'Insert Left',
		iconPath: 'M9 3h8v14H9zM13 3v14M2 10h5M4.5 7.5v5',
		run: (t) => insertNextToSelection(t.element, t.cells, 'left'),
	},
	{
		id: 'tableInsertRight',
		key: 'insertRight',
		label: 'Insert Right',
		iconPath: 'M3 3h8v14H3zM7 3v14M13 10h5M15.5 7.5v5',
		run: (t) => insertNextToSelection(t.element, t.cells, 'right'),
	},
	{
		id: 'tableMergeCells',
		key: 'mergeCells',
		label: 'Merge Cells',
		iconPath: 'M3 4h14v12H3zM3 10h4M13 10h4M7 8l3 2-3 2M13 8l-3 2 3 2',
		large: true,
		enabled: (t) => canMergeSelection(t.data, t.cells),
		run: (t) => mergeSelection(t.element, t.cells),
	},
	{
		id: 'tableSplitCells',
		key: 'splitCells',
		label: 'Split Cells',
		iconPath: 'M3 4h14v12H3zM10 4v12M7 8 4 10l3 2M13 8l3 2-3 2',
		large: true,
		enabled: (t) => canSplitSelection(t.data, t.cells),
		run: (t) => splitSelection(t.element, t.cells),
	},
	{
		id: 'tableDistributeRows',
		key: 'distributeRows',
		label: 'Distribute Rows',
		iconPath: 'M3 3h14M3 10h14M3 17h14M10 4.5v3M10 12.5v3',
		run: (t) => distribute(t.element, t.cells, 'row'),
	},
	{
		id: 'tableDistributeColumns',
		key: 'distributeColumns',
		label: 'Distribute Columns',
		iconPath: 'M3 3v14M10 3v14M17 3v14M4.5 10h3M12.5 10h3',
		run: (t) => distribute(t.element, t.cells, 'column'),
	},
	{
		id: 'tableAlignLeft',
		key: 'alignLeft',
		label: 'Align Left',
		iconPath: 'M3 5h14M3 9h9M3 13h14M3 17h9',
		run: align('left'),
	},
	{
		id: 'tableAlignCenter',
		key: 'alignCenter',
		label: 'Center',
		iconPath: 'M3 5h14M5.5 9h9M3 13h14M5.5 17h9',
		run: align('center'),
	},
	{
		id: 'tableAlignRight',
		key: 'alignRight',
		label: 'Align Right',
		iconPath: 'M3 5h14M8 9h9M3 13h14M8 17h9',
		run: align('right'),
	},
	{
		id: 'tableAlignTop',
		key: 'alignTop',
		label: 'Align Top',
		iconPath: 'M3 3h14M6 6h8v6H6z',
		run: valign('top'),
	},
	{
		id: 'tableAlignMiddle',
		key: 'alignMiddle',
		label: 'Center Vertically',
		iconPath: 'M3 10h2M15 10h2M6 5h8v10H6z',
		run: valign('middle'),
	},
	{
		id: 'tableAlignBottom',
		key: 'alignBottom',
		label: 'Align Bottom',
		iconPath: 'M3 17h14M6 14h8V8H6z',
		run: valign('bottom'),
	},
	{
		id: 'tableCellMargins',
		key: 'cellMargins',
		label: 'Cell Margins',
		iconPath: 'M3 3h14v14H3zM6.5 6.5h7v7h-7z',
		large: true,
		unavailable: IN_INSPECTOR,
	},
	{
		id: 'tableSize',
		key: 'tableSize',
		label: 'Table Size',
		iconPath: 'M4 4h12v12H4zM2 10h2M16 10h2M10 2v2M10 16v2',
		large: true,
		unavailable: IN_INSPECTOR,
	},
	{
		id: 'tableArrange',
		key: 'arrange',
		label: 'Arrange',
		iconPath: 'M3 3h9v9H3zM8 8h9v9H8z',
		large: true,
		unavailable: IN_HOME,
	},
];

function targetsOf(ctx: RibbonGalleryContext): Targets | null {
	const element = ctx.element;
	if (element?.type !== 'table' || !element.tableData) {
		return null;
	}
	return { element, data: element.tableData, cells: ctx.selectedTableCells ?? [] };
}

function moduleFor(spec: CommandSpec): RibbonGalleryModule {
	const labelKey = `pptx.gallery.tableLayout.${spec.key}`;
	return {
		build(ctx) {
			const targets = targetsOf(ctx);
			const gated = targets !== null && spec.enabled !== undefined && !spec.enabled(targets);
			const hint = spec.unavailable ?? (gated ? NEEDS_CELLS : undefined);
			return {
				id: spec.id,
				labelKey,
				label: spec.label,
				disabled: !targets || gated || Boolean(spec.unavailable),
				command: {
					iconPath: spec.iconPath,
					large: spec.large,
					hintKey: hint?.key,
					hint: hint?.text,
				},
				sections: [
					{
						id: 'command',
						columns: 1,
						tileWidth: 0,
						tileHeight: 0,
						items: [{ id: 'run', labelKey, label: spec.label, previewSvg: '', applied: false }],
					},
				],
			};
		},
		apply(itemId, ctx) {
			const targets = targetsOf(ctx);
			if (itemId !== 'run' || !targets || spec.unavailable || !spec.run) {
				return null;
			}
			if (spec.enabled && !spec.enabled(targets)) {
				return null;
			}
			return spec.run(targets);
		},
	};
}

/** Gallery ids that are Table Layout command buttons. */
export type TableLayoutCommandGalleryId = Extract<
	RibbonGalleryId,
	| 'tableSelect'
	| 'tableViewGridlines'
	| 'tableDeleteRows'
	| 'tableDeleteColumns'
	| 'tableInsertAbove'
	| 'tableInsertBelow'
	| 'tableInsertLeft'
	| 'tableInsertRight'
	| 'tableMergeCells'
	| 'tableSplitCells'
	| 'tableDistributeRows'
	| 'tableDistributeColumns'
	| 'tableAlignLeft'
	| 'tableAlignCenter'
	| 'tableAlignRight'
	| 'tableAlignTop'
	| 'tableAlignMiddle'
	| 'tableAlignBottom'
	| 'tableCellMargins'
	| 'tableSize'
	| 'tableArrange'
>;

/** The Table Layout command entries, keyed by gallery id, for the registry. */
export const TABLE_LAYOUT_COMMAND_MODULES = Object.fromEntries(
	SPECS.map((spec) => [spec.id, moduleFor(spec)]),
) as Record<TableLayoutCommandGalleryId, RibbonGalleryModule>;
