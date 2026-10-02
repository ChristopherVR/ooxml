// Clipboard and find/replace payload types of the edit session.
import type { CellRange } from '../address.js';
import type { CellStyle, CellValue, RichTextRun } from '../model.js';

/** One copied cell, self-contained so it can be pasted into another workbook. */
export interface ClipboardCell {
	value: CellValue;
	formula?: string;
	/** The resolved format (not an index), so it survives a paste into another workbook. */
	style?: CellStyle;
	richText?: RichTextRun[];
	/** Number format implied by pasted text (`15%`), applied when the target is `General`. */
	numFmt?: string;
	/** Display text as the grid shows it. */
	text: string;
	/** The formula uses pre-dynamic-array semantics (see `Cell.legacyFormula`). */
	legacyFormula?: true;
	/**
	 * A dynamic-array result whose anchor formula is part of the same copy: pasting formulas
	 * recreates it from the anchor, so only `values` pastes write it.
	 */
	spilled?: true;
}

export interface ClipboardCells {
	rows: number;
	cols: number;
	/** `data[r][c]`; `null` for an empty cell. */
	data: (ClipboardCell | null)[][];
	/** Merged areas, relative to the top-left of the copied block. */
	merges: CellRange[];
	/** Where the cells were copied from, when they came from a session. */
	source?: { sheet: number; range: CellRange };
}

/** What `copy` returns: text for the system clipboard plus the cells for an internal paste. */
export interface ClipboardPayload {
	tsv: string;
	html: string;
	cells: ClipboardCells;
	/** Set by `cut`; the first paste clears the source and then resets this flag. */
	cut?: boolean;
}

export interface FindQuery {
	text: string;
	/** Sheet to search; all sheets when absent. */
	sheet?: number;
	/** Limit the search to a range (only with `sheet`). */
	range?: CellRange;
	matchCase?: boolean;
	/** The whole cell content must match. */
	wholeCell?: boolean;
	/** `formulas` searches what was typed (formula text, raw constants); `values` the display text. */
	lookIn?: 'formulas' | 'values' | 'comments';
	/** Excel wildcards: `*` any run, `?` one character, `~` escapes. Default true. */
	wildcards?: boolean;
	/** Search order; rows by default. */
	order?: 'rows' | 'columns';
}

export interface FindMatch {
	sheet: number;
	row: number;
	col: number;
	/** The text that was searched in this cell. */
	text: string;
}
