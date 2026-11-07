/** Shared logical table edits for the editor and the PPTX automation tools. */
import type { PptxTableCell, PptxTableData } from '../../core/types';
import type { TableStructureEdit } from '../../core/core/runtime/table-structure-edit-xml';

/** Default pixel height assigned to a freshly inserted row. */
export const DEFAULT_INSERTED_ROW_HEIGHT = 40;

function createCell(text = ''): PptxTableCell {
	return { text, style: {} };
}

function copyRows(data: PptxTableData): PptxTableData['rows'] {
	return data.rows.map((row) => ({ ...row, cells: [...row.cells] }));
}

/**
 * Rebase the merges that cross the edited line using the original anchors.
 * The caller owns the new rows and cell arrays. Cells outside these merges
 * already moved with the splice, so their contents and metadata can stay intact.
 */
function updateIntersectingMerges(
	source: PptxTableData,
	rows: PptxTableData['rows'],
	edit: TableStructureEdit,
): void {
	const isRowEdit = edit.axis === 'row';
	const inserting = edit.action === 'insert';
	const delta = inserting ? 1 : -1;
	for (const [rowIndex, row] of source.rows.entries()) {
		for (const [columnIndex, anchor] of row.cells.entries()) {
			if (anchor.hMerge || anchor.vMerge) continue;
			const rowSpan = anchor.rowSpan ?? 1;
			const gridSpan = anchor.gridSpan ?? 1;
			const start = isRowEdit ? rowIndex : columnIndex;
			const span = isRowEdit ? rowSpan : gridSpan;
			if (
				edit.index < start ||
				edit.index >= start + span ||
				(inserting ? edit.index === start : span === 1)
			) {
				continue;
			}
			if (
				!Number.isInteger(rowSpan) ||
				!Number.isInteger(gridSpan) ||
				rowSpan < 1 ||
				gridSpan < 1 ||
				rowIndex + rowSpan > source.rows.length ||
				columnIndex + gridSpan > source.columnWidths.length
			) {
				continue;
			}

			const height = isRowEdit ? rowSpan + delta : rowSpan;
			const width = isRowEdit ? gridSpan : gridSpan + delta;
			const anchorRemoved = !inserting && edit.index === start;
			for (let r = rowIndex; r < rowIndex + height; r++) {
				for (let c = columnIndex; c < columnIndex + width; c++) {
					const cell = rows[r]?.cells[c];
					if (!cell) continue;
					const firstRow = r === rowIndex;
					const firstColumn = c === columnIndex;
					const promote = anchorRemoved && firstRow && firstColumn;
					const content = promote && anchor.text ? anchor : cell;
					// DrawingML repeats the height along the top edge and the width
					// along the left edge of a rectangular merge (MS-OE376, tbl).
					rows[r].cells[c] = {
						...cell,
						text: content.text,
						textRuns: content.textRuns,
						paragraphs: content.paragraphs,
						style: promote ? (cell.style ?? anchor.style) : cell.style,
						rowSpan: firstRow && height > 1 ? height : undefined,
						gridSpan: firstColumn && width > 1 ? width : undefined,
						hMerge: firstColumn ? undefined : true,
						vMerge: firstRow ? undefined : true,
					};
				}
			}
		}
	}
}

function normalizeWidths(widths: number[]): number[] {
	const sum = widths.reduce((total, width) => total + width, 0);
	return sum > 0 ? widths.map((width) => width / sum) : widths;
}

/** Insert a row and extend each vertical merge crossing the insertion point. */
export function insertTableRow(
	tableData: PptxTableData,
	rowIdx: number,
	position: 'above' | 'below',
	cellTexts?: readonly string[],
): PptxTableData {
	const index = Math.max(0, Math.min(tableData.rows.length, rowIdx + Number(position === 'below')));
	const rows = copyRows(tableData);
	rows.splice(index, 0, {
		height: DEFAULT_INSERTED_ROW_HEIGHT,
		cells: tableData.columnWidths.map((_, column) => createCell(cellTexts?.[column])),
	});
	updateIntersectingMerges(tableData, rows, { axis: 'row', action: 'insert', index });
	return { ...tableData, rows };
}

/** Delete a row, retaining the title and formatting of any surviving merge. */
export function deleteTableRow(tableData: PptxTableData, rowIdx: number): PptxTableData {
	if (tableData.rows.length <= 1 || rowIdx < 0 || rowIdx >= tableData.rows.length) {
		return tableData;
	}
	const rows = copyRows(tableData);
	rows.splice(rowIdx, 1);
	updateIntersectingMerges(tableData, rows, { axis: 'row', action: 'delete', index: rowIdx });
	return { ...tableData, rows };
}

/** Insert a column, splitting a neighbour's width and extending intersected merges. */
export function insertTableColumn(
	tableData: PptxTableData,
	colIdx: number,
	position: 'left' | 'right',
	cellTexts?: readonly string[],
): PptxTableData {
	const count = tableData.columnWidths.length;
	const index = Math.max(0, Math.min(count, colIdx + Number(position === 'right')));
	const columnWidths = [...tableData.columnWidths];
	if (count === 0) {
		columnWidths.push(1);
	} else {
		const neighbour = Math.min(index, count - 1);
		const half = columnWidths[neighbour] / 2;
		columnWidths[neighbour] = half;
		columnWidths.splice(index, 0, half);
	}
	const rows = copyRows(tableData);
	rows.forEach((row, rowIndex) => row.cells.splice(index, 0, createCell(cellTexts?.[rowIndex])));
	updateIntersectingMerges(tableData, rows, { axis: 'column', action: 'insert', index });
	return { ...tableData, rows, columnWidths: normalizeWidths(columnWidths) };
}

/** Delete a column and move any surviving merge's anchor onto the remaining grid. */
export function deleteTableColumn(tableData: PptxTableData, colIdx: number): PptxTableData {
	if (tableData.columnWidths.length <= 1 || colIdx < 0 || colIdx >= tableData.columnWidths.length) {
		return tableData;
	}
	const rows = copyRows(tableData);
	rows.forEach((row) => row.cells.splice(colIdx, 1));
	updateIntersectingMerges(tableData, rows, { axis: 'column', action: 'delete', index: colIdx });
	const columnWidths = tableData.columnWidths.filter((_, index) => index !== colIdx);
	return { ...tableData, rows, columnWidths: normalizeWidths(columnWidths) };
}
