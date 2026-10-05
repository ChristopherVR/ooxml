// Dynamic-array formulas on save: which formula cells are written as `t="array"` anchors with
// XLDAPR cell metadata, the spill range each records, and the `xl/metadata.xml` part.
import { type CellRange, cellKey } from '../address.js';
import { needsArrayEvaluation } from '../formula/array-context.js';
import { isSpilledCell, type SpilledCell } from '../formula/spill.js';
import { type Cell, isCellError, type Worksheet } from '../model.js';
import { XML_HEADER } from './xml-out.js';

/** Spill footprints of a sheet, computed from the spilled cells the calc engine left. */
export class SpillPlan {
	private readonly ranges = new Map<number, CellRange>();

	constructor(private readonly sheet: Worksheet) {
		for (const [row, cells] of sheet.rows)
			for (const [col, cell] of cells) {
				if (!isSpilledCell(cell)) continue;
				const { row: ar, col: ac } = cell.spillAnchor;
				const key = cellKey(ar, ac);
				const range = this.ranges.get(key) ?? {
					start: { row: ar, col: ac },
					end: { row: ar, col: ac },
				};
				range.start = { row: Math.min(range.start.row, row), col: Math.min(range.start.col, col) };
				range.end = { row: Math.max(range.end.row, row), col: Math.max(range.end.col, col) };
				this.ranges.set(key, range);
			}
	}

	/** The range to record for a dynamic-array anchor, or undefined for an ordinary formula. */
	dynamicRange(cell: Cell, row: number, col: number): CellRange | undefined {
		if (!cell.formula || cell.arrayRange || cell.legacyFormula) return undefined;
		const spill = this.ranges.get(cellKey(row, col));
		if (spill && spill.start.row === row && spill.start.col === col) return spill;
		const blocked = isCellError(cell.value) && cell.value.error === '#SPILL!';
		if (cell.dynamicArray || blocked || needsArrayEvaluation(cell.formula))
			return { start: { row, col }, end: { row, col } };
		return undefined;
	}

	/** Whether a spilled cell lies in the range its (still present) anchor will record. */
	covered(cell: SpilledCell, row: number, col: number): boolean {
		const { row: ar, col: ac } = cell.spillAnchor;
		const anchor = this.sheet.rows.get(ar)?.get(ac);
		if (!anchor?.formula || anchor.arrayRange || anchor.legacyFormula) return false;
		const range = this.dynamicRange(anchor, ar, ac);
		return (
			!!range &&
			row >= range.start.row &&
			row <= range.end.row &&
			col >= range.start.col &&
			col <= range.end.col
		);
	}
}

/** `xl/metadata.xml` declaring cell metadata block 1 as an Excel dynamic array (XLDAPR). */
export const DYNAMIC_ARRAY_METADATA_XML =
	`${XML_HEADER}<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
	'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
	'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" ' +
	'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" ' +
	'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
	'<futureMetadata name="XLDAPR" count="1"><bk><extLst>' +
	'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">' +
	'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
	'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
