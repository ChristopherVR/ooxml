/**
 * The inspector's table options (PowerPoint's Table Design > Table Style
 * Options, plus the band sizes, the table direction and a uniform cell
 * margin), as field descriptors every binding draws with its one generic
 * renderer.
 *
 * The descriptors read the element's `PptxTableData` and write a
 * `Partial<PptxTableData>` the binding merges into it through its usual table
 * route. The compound table controls (the style preset gallery, the style
 * editor, the structure buttons and the per-cell formatting) stay native
 * templates beside them.
 *
 * @module render/inspector/table-option-fields
 */
import type { PptxTableData } from 'ooxml-core/pptx';

import type { InspectorFieldDescriptor, InspectorSectionDescriptor } from './inspector-field-model';

export type TableOptionsTarget = PptxTableData | undefined;
export type TableOptionsSection = InspectorSectionDescriptor<
	TableOptionsTarget,
	Partial<PptxTableData>
>;
type TableOptionsField = InspectorFieldDescriptor<TableOptionsTarget, Partial<PptxTableData>>;

type TableFlag =
	| 'firstRowHeader'
	| 'firstCol'
	| 'lastRow'
	| 'lastCol'
	| 'bandedRows'
	| 'bandedColumns'
	| 'rtl';

/** Points per CSS pixel (72 pt and 96 px to the inch). */
const PT_PER_PX = 0.75;
/** Largest band size the fields offer. */
const MAX_BAND_SIZE = 99;

function flag(id: TableFlag, labelKey: string): TableOptionsField {
	return {
		id,
		kind: 'toggle',
		labelKey,
		testId: `table-${id}`,
		read: (table) => table?.[id] === true,
		write: (_table, value) => ({ [id]: value === true }),
	};
}

function bandSize(
	id: 'bandRowCycle' | 'bandColCycle',
	when: 'bandedRows' | 'bandedColumns',
	labelKey: string,
): TableOptionsField {
	return {
		id,
		kind: 'number',
		labelKey,
		min: 1,
		max: MAX_BAND_SIZE,
		step: 1,
		decimals: 0,
		testId: `table-${id}`,
		visibleWhen: (table) => table?.[when] === true,
		read: (table) => table?.[id] ?? 1,
		write: (_table, value) => ({ [id]: Math.round(Number(value)) }),
	};
}

/** The cell margin the fields show: the first cell's left margin, in points. */
function cellMarginOf(table: TableOptionsTarget): number {
	return (table?.rows[0]?.cells[0]?.style?.marginLeft ?? 0) * PT_PER_PX;
}

/** The same margin on every side of every cell (`a:tcPr/@marL` ... `@marB`). */
function uniformCellMargin(table: TableOptionsTarget, points: number): Partial<PptxTableData> {
	if (!table) {
		return {};
	}
	const margin = Math.round(points / PT_PER_PX);
	return {
		rows: table.rows.map((row) => ({
			...row,
			cells: row.cells.map((cell) => ({
				...cell,
				style: {
					...cell.style,
					marginLeft: margin,
					marginRight: margin,
					marginTop: margin,
					marginBottom: margin,
				},
			})),
		})),
	};
}

/** Table Style Options, the band sizes, the direction and the cell margin (untitled). */
export const TABLE_OPTIONS_SECTION: TableOptionsSection = {
	id: 'tableOptions',
	fields: [
		flag('firstRowHeader', 'pptx.table.headerRow'),
		flag('firstCol', 'pptx.table.firstColumn'),
		flag('lastRow', 'pptx.table.totalRow'),
		flag('lastCol', 'pptx.table.lastColumn'),
		flag('bandedRows', 'pptx.table.bandedRows'),
		flag('bandedColumns', 'pptx.table.bandedColumns'),
		bandSize('bandRowCycle', 'bandedRows', 'pptx.table.bandRowCycle'),
		bandSize('bandColCycle', 'bandedColumns', 'pptx.table.bandColCycle'),
		flag('rtl', 'pptx.textAdvanced.rtl'),
		{
			id: 'cellMargin',
			kind: 'length',
			labelKey: 'pptx.table.cellPadding',
			min: 0,
			max: 288,
			step: 1,
			unit: 'pt',
			decimals: 1,
			wide: true,
			testId: 'table-cellMargin',
			read: cellMarginOf,
			write: (table, value) => uniformCellMargin(table, Number(value)),
		},
	],
};

/** The table options as a section list, what a generic renderer takes. */
export const TABLE_OPTIONS_SECTIONS: readonly TableOptionsSection[] = [TABLE_OPTIONS_SECTION];
