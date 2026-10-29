import { flatCellParagraphs, flatCells, flatRows } from './parse-content-control';
import { expectDefined } from './expect-defined';
import type { Block, Paragraph, Table, TableCell } from './model';
import {
	patchCellBorders,
	patchCellShading,
	patchCellMargins,
	patchCellVerticalAlignment,
	patchRowProperties,
	patchTableMargins,
} from './write-table-properties';
import { placeTableControl, unwrapDroppedControls } from './write-table-controls';
import { syncBlockControls } from './write-block-placement';
import { children, first, makeW, type XmlDocument, type XmlElement } from './xml';

type ParagraphWriter = (
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base?: Paragraph,
) => XmlElement;
type SlotsWriter = (
	doc: XmlDocument,
	parent: XmlElement,
	old: XmlElement[],
	next: XmlElement[],
) => void;

import { canEditTableStructure, canRestructureTable } from './table-structure-check';
import { restructured, writeRestructuredTable } from './write-table-structure';
import { tableGrid } from './table-grid';

export { canEditTableStructure };

const layout = (table: Table) =>
	table.rows.map((row) => row.map((cell) => cell.paragraphs.map((p) => p.id)));

/** Table/cell descriptive fields this writer does not yet serialize back to XML. */
const TABLE_DESCRIPTOR_KEYS = [
	'grid',
	'widthTwips',
	'alignment',
	'indentTwips',
	'borders',
	'style',
	'look',
] as const;
const CELL_DESCRIPTOR_KEYS = ['gridSpan', 'verticalMerge', 'widthTwips'] as const;
function pick(source: object, keys: readonly string[]): Record<string, unknown> {
	const record = source as Record<string, unknown>;
	const result: Record<string, unknown> = {};
	for (const key of keys) if (record[key] !== undefined) result[key] = record[key];
	return result;
}
/**
 * Table grid/width/border/style and per-cell width/merge values render but cannot be edited yet.
 * Cell borders, shading, margins and vertical alignment have preservation-safe writers below.
 */
function assertNoDescriptorEdits(table: Table, base: Table | undefined, layoutEdits = false): void {
	if (!base) return;
	const tableKeys = layoutEdits
		? TABLE_DESCRIPTOR_KEYS.filter((key) => key !== 'grid')
		: TABLE_DESCRIPTOR_KEYS;
	if (JSON.stringify(pick(table, tableKeys)) !== JSON.stringify(pick(base, tableKeys)))
		throw new Error(
			'Cannot edit table grid widths, width, alignment, indent, borders, or style; only cell text and row/column structure changes are supported. The original DOCX package remains unchanged.',
		);
	const baseCellById = new Map<string, TableCell>();
	if (layoutEdits) return;
	for (const row of base.rows)
		for (const cell of row) {
			const id = cell.paragraphs[0]?.id;
			if (id) baseCellById.set(id, cell);
		}
	for (const row of table.rows)
		for (const cell of row) {
			const id = cell.paragraphs[0]?.id;
			const source = id ? baseCellById.get(id) : undefined;
			if (!source) continue;
			if (
				JSON.stringify(pick(cell, CELL_DESCRIPTOR_KEYS)) !==
				JSON.stringify(pick(source, CELL_DESCRIPTOR_KEYS))
			)
				throw new Error(
					'Cannot edit table cell width or merge on an existing cell; cell text, borders, shading, margins and vertical alignment are supported. The original DOCX package remains unchanged.',
				);
		}
}

/** Patch edited paragraphs of tables nested in a cell back into their untouched source XML. */
function writeNestedTables(
	doc: XmlDocument,
	tc: XmlElement,
	cell: TableCell,
	base: TableCell | undefined,
	writeParagraph: ParagraphWriter,
	replaceSlots: SlotsWriter,
): void {
	const tables = (blocks: Block[] | undefined) =>
		(blocks ?? []).filter((block): block is Table => block.type === 'table');
	const next = tables(cell.sourceBlocks);
	const before = tables(base?.sourceBlocks);
	const nodes = children(tc, 'tbl');
	next.forEach((table, index) => {
		const node = nodes[index];
		if (node && before[index])
			writeTable(doc, table, node, before[index], writeParagraph, replaceSlots, true);
	});
}

export function writeTable(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base: Table | undefined,
	writeParagraph: ParagraphWriter,
	replaceSlots: SlotsWriter,
	nested = false,
): XmlElement {
	if (base && JSON.stringify(table) === JSON.stringify(base)) return node;
	// Merged tables, tables nested in a cell and edits that merge or split cells restructure
	// through the merge-aware writer; plain rectangular tables keep the row and column writer below.
	const baseGrid = base && tableGrid(base);
	if (
		base &&
		baseGrid &&
		(nested || !base.structureEditable || table.structureEditable === false) &&
		restructured(table, base) &&
		canRestructureTable(node)
	) {
		assertNoDescriptorEdits(table, base, true);
		return writeRestructuredTable(
			doc,
			table,
			node,
			base,
			writeParagraph,
			replaceSlots,
			(d, tc, cell, before) => writeNestedTables(d, tc, cell, before, writeParagraph, replaceSlots),
		);
	}
	assertNoDescriptorEdits(table, base);
	patchTableMargins(doc, table, node, base);
	const structural = base && JSON.stringify(layout(table)) !== JSON.stringify(layout(base));
	if (structural && !canEditTableStructure(node))
		throw new Error(
			'Cannot change the structure of a merged, nested, or complex table. The original DOCX package remains unchanged.',
		);
	const width = table.rows[0]?.length ?? 0;
	// The rows of a merged table differ in physical cell count (a label beside a value spanning
	// three columns). A text edit leaves every cell where it is, so only a structural edit, which
	// canEditTableStructure limits to plain grids, needs rows of a single width.
	const reshaped = structural || !base;
	if (
		!width ||
		table.rows.some(
			(row) => (reshaped && row.length !== width) || row.some((cell) => !cell.paragraphs.length),
		)
	)
		throw new Error('Tables must contain rectangular rows and at least one paragraph per cell.');
	const oldRows = flatRows(node).map((entry) => entry.element);
	type Source = {
		row: number;
		col: number;
		cell: XmlElement;
		paragraphs: Paragraph[];
		model: TableCell;
	};
	const sourceById = new Map<string, Source>();
	base?.rows.forEach((row, ri) =>
		row.forEach((cell, ci) => {
			const oldCell = flatCells(expectDefined(oldRows[ri], `table row ${ri}`))[ci]?.element;
			if (oldCell)
				for (const paragraph of cell.paragraphs)
					sourceById.set(paragraph.id, {
						row: ri,
						col: ci,
						cell: oldCell,
						paragraphs: cell.paragraphs,
						model: cell,
					});
		}),
	);
	const usedCells = new Set<XmlElement>();
	const usedRows = new Set<number>();
	const columns = new Map<number, number>();
	const rowSources: (number | undefined)[] = [];
	const nextRows = table.rows.map((row, rowIndex) => {
		const sources = row.map((cell) =>
			cell.paragraphs.map((p) => sourceById.get(p.id)).find(Boolean),
		);
		const sourceRow = sources.find(Boolean)?.row;
		if (sources.some((source) => source && source.row !== sourceRow))
			throw new Error('Moving paragraphs between table rows is not supported.');
		if (sourceRow !== undefined && usedRows.has(sourceRow))
			throw new Error('Splitting an existing table row is not supported.');
		if (sourceRow !== undefined) usedRows.add(sourceRow);
		rowSources.push(sourceRow);
		const tr =
			sourceRow === undefined
				? makeW(doc, 'tr')
				: expectDefined(oldRows[sourceRow], `table row ${sourceRow}`);
		const oldCells = flatCells(tr).map((entry) => entry.element);
		patchRowProperties(
			doc,
			tr,
			table.rowProperties?.[rowIndex],
			sourceRow === undefined ? undefined : base?.rowProperties?.[sourceRow],
		);
		const nextCells = row.map((cell, ci) => {
			const source = sources[ci];
			if (source && usedCells.has(source.cell))
				throw new Error('Splitting an existing table cell is not supported.');
			if (source) {
				usedCells.add(source.cell);
				if (columns.has(ci) && columns.get(ci) !== source.col)
					throw new Error('Inconsistent table column edits are not supported.');
				columns.set(ci, source.col);
			}
			if (
				cell.paragraphs.some(
					(p) => sourceById.has(p.id) && sourceById.get(p.id)?.cell !== source?.cell,
				)
			)
				throw new Error('Moving paragraphs between table cells is not supported.');
			const tc = source?.cell ?? makeW(doc, 'tc');
			patchCellBorders(doc, tc, cell.borders, source?.model.borders);
			patchCellShading(doc, tc, cell, source?.model);
			patchCellMargins(doc, tc, cell.margins, source?.model.margins);
			patchCellVerticalAlignment(doc, tc, cell.verticalAlign, source?.model.verticalAlign);
			const flatParagraphs = flatCellParagraphs(tc);
			const oldParagraphs = flatParagraphs.map((entry) => entry.element);
			const nextParagraphs = cell.paragraphs.map((paragraph) => {
				const index = source?.paragraphs.findIndex((p) => p.id === paragraph.id) ?? -1;
				return writeParagraph(
					doc,
					paragraph,
					oldParagraphs[index] ?? makeW(doc, 'p'),
					source?.paragraphs[index],
				);
			});
			replaceSlots(doc, tc, oldParagraphs, nextParagraphs);
			// A block control around a cell's paragraphs follows the model as in the body.
			syncBlockControls(doc, flatParagraphs, cell.paragraphs, nextParagraphs);
			writeNestedTables(doc, tc, cell, source?.model, writeParagraph, replaceSlots);
			return tc;
		});
		replaceSlots(doc, tr, oldCells, nextCells);
		nextCells.forEach((tc, ci) =>
			placeTableControl(doc, tc, row[ci]?.contentControl, sources[ci]?.model.contentControl),
		);
		unwrapDroppedControls(
			nextCells,
			row.map((cell) => cell.contentControl),
		);
		return tr;
	});
	replaceSlots(doc, node, oldRows, nextRows);
	nextRows.forEach((tr, ri) => {
		const sourceRow = rowSources[ri];
		placeTableControl(
			doc,
			tr,
			table.rowProperties?.[ri]?.contentControl,
			sourceRow === undefined ? undefined : base?.rowProperties?.[sourceRow]?.contentControl,
		);
	});
	unwrapDroppedControls(
		nextRows,
		nextRows.map((_, ri) => table.rowProperties?.[ri]?.contentControl),
	);
	const grid = first(node, 'tblGrid');
	if (
		structural &&
		grid &&
		(width !== base.rows[0]?.length || [...columns].some(([next, old]) => next !== old))
	) {
		const oldColumns = children(grid, 'gridCol');
		const nextColumns = Array.from({ length: width }, (_, index) => {
			const old = oldColumns[columns.get(index) ?? -1];
			return (
				old ??
				(oldColumns[Math.min(index, oldColumns.length - 1)]?.cloneNode(true) as
					| XmlElement
					| undefined) ??
				makeW(doc, 'gridCol')
			);
		});
		replaceSlots(doc, grid, oldColumns, nextColumns);
	}
	return node;
}
