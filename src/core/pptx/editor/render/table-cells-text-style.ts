/**
 * table-cells-text-style.ts: character formatting applied to a set of table
 * cells (a cell range, or the whole table), as PowerPoint does when Bold,
 * a colour or a size is picked with cells or the table frame selected: every
 * run in every selected cell takes it, and an empty cell keeps it for the
 * text typed into it later.
 *
 * Three representations change together: the cell style (the empty-cell and
 * single-run fallback), the runs a renderer paints (`textRuns`, and
 * `textSegments` for a cell edited this session), and the cell's markup in
 * `rawXml`, where every `a:rPr` (or the `a:endParaRPr` of an empty
 * paragraph) is updated in place so the rest of each run survives verbatim.
 *
 * @module render/table-cells-text-style
 */
import type {
	PptxTableCell,
	PptxTableCellStyle,
	PptxTableCellTextRun,
	TablePptxElement,
	TextStyle,
} from 'ooxml-core/pptx';
import { updateCellTextStyleInRawXml } from 'ooxml-core/pptx';

import { withCellStyle } from './table-cell-edit';
import { tableCellSegments, withCellSegments } from './table-cell-rich-text';
import { updateTextSegmentStyle } from './update-text-segment-style';

/** A cell position in grid coordinates. */
export interface TableCellRef {
	row: number;
	col: number;
}

/**
 * Character (and alignment) formatting for table cells, in cell units: sizes
 * in points, as on `PptxTableCellStyle`.
 */
export type TableCellTextStyleUpdate = Partial<
	Pick<
		PptxTableCellStyle,
		'bold' | 'italic' | 'underline' | 'color' | 'fontFamily' | 'fontSize' | 'align'
	>
> & { strikethrough?: boolean };

const RUN_KEYS = [
	'bold',
	'italic',
	'underline',
	'strikethrough',
	'color',
	'fontFamily',
	'fontSize',
] as const;

/** The cells a selection covers: the given ones, or every anchor cell of the table. */
function targetCells(element: TablePptxElement, cells?: readonly TableCellRef[]): TableCellRef[] {
	if (cells && cells.length > 0) {
		return [...cells];
	}
	const all: TableCellRef[] = [];
	element.tableData?.rows.forEach((row, rowIndex) =>
		row.cells.forEach((cell, colIndex) => {
			if (!cell.hMerge && !cell.vMerge) {
				all.push({ row: rowIndex, col: colIndex });
			}
		}),
	);
	return all;
}

function runPatch(update: TableCellTextStyleUpdate): Partial<PptxTableCellTextRun> {
	const patch: Partial<PptxTableCellTextRun> = {};
	for (const key of RUN_KEYS) {
		if (key in update) {
			(patch as Record<string, unknown>)[key] = update[key];
		}
	}
	return patch;
}

function segmentPatch(update: TableCellTextStyleUpdate): Partial<TextStyle> {
	const patch = runPatch(update) as Partial<TextStyle>;
	if (typeof update.fontSize === 'number') {
		patch.fontSize = update.fontSize * (96 / 72);
	}
	return patch;
}

function cellStylePatch(update: TableCellTextStyleUpdate): Partial<PptxTableCellStyle> {
	const { strikethrough: _strikethrough, ...patch } = update;
	return patch;
}

/** One cell with the formatting applied to its style and every run. */
export function withCellTextStyle(
	cell: PptxTableCell,
	update: TableCellTextStyleUpdate,
): PptxTableCell {
	const styled = withCellStyle(cell, cellStylePatch(update));
	if (styled.textSegments?.length) {
		const patch = segmentPatch(update);
		const segments = tableCellSegments(styled).map((segment) => {
			if (segment.isParagraphBreak) {
				return segment;
			}
			const next = updateTextSegmentStyle(segment, patch);
			if (update.align && next.paragraphProperties) {
				next.paragraphProperties = { ...next.paragraphProperties, align: update.align };
			}
			return next;
		});
		return withCellSegments(styled, segments);
	}
	if (styled.textRuns?.length) {
		const patch = runPatch(update);
		styled.textRuns = styled.textRuns.map((run) =>
			run.isParagraphBreak || run.isLineBreak ? run : { ...run, ...patch },
		);
	}
	return styled;
}

/**
 * Apply character formatting to every run of the given cells (every cell of
 * the table when `cells` is empty or omitted), keeping `rawXml` in step.
 * Returns the element itself when it has no table data.
 */
export function applyTableCellsTextStyle(
	element: TablePptxElement,
	cells: readonly TableCellRef[] | undefined,
	update: TableCellTextStyleUpdate,
): TablePptxElement {
	const tableData = element.tableData;
	if (!tableData) {
		return element;
	}
	const targets = targetCells(element, cells);
	const keys = new Set(targets.map(({ row, col }) => `${row}:${col}`));
	const rows = tableData.rows.map((row, rowIndex) => ({
		...row,
		cells: row.cells.map((cell, colIndex) =>
			keys.has(`${rowIndex}:${colIndex}`) ? withCellTextStyle(cell, update) : cell,
		),
	}));
	const next: TablePptxElement = { ...element, tableData: { ...tableData, rows } };
	let rawXml = element.rawXml;
	for (const { row, col } of targets) {
		if (!rawXml) {
			break;
		}
		rawXml =
			updateCellTextStyleInRawXml(
				{ ...element, rawXml },
				row,
				col,
				update as Record<string, unknown>,
			) ?? rawXml;
	}
	if (rawXml) {
		next.rawXml = rawXml;
	}
	return next;
}
