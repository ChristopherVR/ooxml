/**
 * table-tsv-paste.ts: pasting tab-separated plain text (cells copied from a
 * spreadsheet, or typed with tabs) into a table. The text is split into a
 * grid and written from the anchor cell, one value per cell, adding rows
 * and columns when the grid runs past the table's edge.
 *
 * @module render/table-tsv-paste
 */
import type { PptxTableData } from 'ooxml-core/pptx';
import type { TableGridFill, TableGridWrite } from 'ooxml-core/pptx/editor/render/table-grid-fill';

/**
 * Split clipboard text into rows of cell values: rows on line breaks (CRLF,
 * CR or LF), values on tabs. A value in double quotes may hold tabs, line
 * breaks and doubled quotes, as spreadsheets write them. One trailing line
 * break (spreadsheets end every copied row with one) adds no row.
 */
export function parseTabularText(text: string): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let value = '';
	let index = 0;
	let quoted = false;
	const normalized = text.replace(/\r\n?/g, '\n');
	while (index < normalized.length) {
		const ch = normalized[index]!;
		if (quoted) {
			if (ch === '"' && normalized[index + 1] === '"') {
				value += '"';
				index += 2;
				continue;
			}
			if (ch === '"') {
				quoted = false;
				index++;
				continue;
			}
			value += ch;
			index++;
			continue;
		}
		if (ch === '"' && value === '') {
			quoted = true;
		} else if (ch === '\t') {
			row.push(value);
			value = '';
		} else if (ch === '\n') {
			row.push(value);
			rows.push(row);
			row = [];
			value = '';
		} else {
			value += ch;
		}
		index++;
	}
	if (value !== '' || row.length > 0 || rows.length === 0) {
		row.push(value);
		rows.push(row);
	}
	return rows;
}

/**
 * Whether pasted text should fill cells rather than go into the edited cell
 * as text: it holds a tab, or several lines that end with a line break (a
 * column of spreadsheet cells). Plain multi-line text without either stays
 * in one cell as paragraphs, as PowerPoint pastes it.
 */
export function isTabularPaste(text: string): boolean {
	if (text.includes('\t')) {
		return true;
	}
	const normalized = text.replace(/\r\n?/g, '\n');
	return normalized.endsWith('\n') && normalized.slice(0, -1).includes('\n');
}

/** A cell position in grid coordinates. */
export interface TablePasteAnchor {
	row: number;
	col: number;
}

/**
 * Plan a grid paste at `anchor`: how many rows and columns the table must
 * grow by, and which value goes into which cell. Positions inside a merged
 * cell are left out, so a merge keeps the value written to its owner.
 */
export function planTableGridPaste(
	data: PptxTableData,
	anchor: TablePasteAnchor,
	grid: readonly (readonly string[])[],
): TableGridFill {
	const width = grid.reduce((max, row) => Math.max(max, row.length), 0);
	const rowCount = data.rows.length;
	const colCount = data.columnWidths.length;
	const writes: TableGridWrite[] = [];
	grid.forEach((values, r) =>
		values.forEach((text, c) => {
			const row = anchor.row + r;
			const col = anchor.col + c;
			const cell = data.rows[row]?.cells[col];
			if (cell?.hMerge || cell?.vMerge) {
				return;
			}
			writes.push({ row, col, text });
		}),
	);
	return {
		addRows: Math.max(0, anchor.row + grid.length - rowCount),
		addColumns: Math.max(0, anchor.col + width - colCount),
		writes,
	};
}
