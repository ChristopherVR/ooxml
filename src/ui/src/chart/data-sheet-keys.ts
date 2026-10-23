/**
 * Keyboard behaviour of the chart data sheet, as a pure decision so it is tested without a DOM.
 * It follows the spreadsheet PowerPoint opens for Edit Data:
 *
 * - a focused cell is `ready`: arrows move to the neighbouring cell, typing replaces the text
 *   and enters `enter` mode, F2 enters `edit` mode with the caret at the end;
 * - in `enter` mode arrows commit and move, in `edit` mode they move the caret;
 * - Enter commits and moves down (Shift+Enter up), Tab right (Shift+Tab left); Tab past the last
 *   column and Shift+Tab before the first leave the sheet, as they would leave any field;
 * - Escape restores the cell's text and returns to `ready`.
 *
 * Cells use the grid coordinates of `ooxml-core/chart` `data-sheet`: row 0 is the series names,
 * column 0 the category labels, and the corner cell (0, 0) does not exist.
 */

export type ChartDataSheetMode = 'ready' | 'enter' | 'edit';

export interface ChartDataSheetKey {
	key: string;
	shiftKey?: boolean;
	ctrlKey?: boolean;
	metaKey?: boolean;
	altKey?: boolean;
}

export type ChartDataSheetKeyAction =
	| { type: 'move'; row: number; col: number }
	| { type: 'edit' }
	| { type: 'revert' }
	| { type: 'typing' }
	| { type: 'native' };

const STEPS: Readonly<Record<string, readonly [number, number]>> = {
	ArrowUp: [-1, 0],
	ArrowDown: [1, 0],
	ArrowLeft: [0, -1],
	ArrowRight: [0, 1],
};

function exists(row: number, col: number, rows: number, cols: number): boolean {
	return row >= 0 && col >= 0 && row < rows && col < cols && !(row === 0 && col === 0);
}

/** Where a step from a cell lands: the cell itself when the step would leave the grid. */
function step(
	row: number,
	col: number,
	[dr, dc]: readonly [number, number],
	rows: number,
	cols: number,
): { row: number; col: number } {
	return exists(row + dr, col + dc, rows, cols) ? { row: row + dr, col: col + dc } : { row, col };
}

/** What a key does in a cell at (`row`, `col`) of a sheet `rows` x `cols` (corner included). */
export function chartDataSheetKeyAction(
	key: ChartDataSheetKey,
	mode: ChartDataSheetMode,
	row: number,
	col: number,
	rows: number,
	cols: number,
): ChartDataSheetKeyAction {
	if (key.ctrlKey || key.metaKey || key.altKey) return { type: 'native' };
	const arrow = STEPS[key.key];
	if (arrow) {
		return mode === 'edit'
			? { type: 'native' }
			: { type: 'move', ...step(row, col, arrow, rows, cols) };
	}
	switch (key.key) {
		case 'F2':
			return { type: 'edit' };
		case 'Escape':
			return { type: 'revert' };
		case 'Enter':
			return { type: 'move', ...step(row, col, key.shiftKey ? [-1, 0] : [1, 0], rows, cols) };
		case 'Tab': {
			const target = step(row, col, key.shiftKey ? [0, -1] : [0, 1], rows, cols);
			return target.col === col ? { type: 'native' } : { type: 'move', ...target };
		}
		case 'Backspace':
		case 'Delete':
			return mode === 'ready' ? { type: 'typing' } : { type: 'native' };
		default:
			return key.key.length === 1 && mode === 'ready' ? { type: 'typing' } : { type: 'native' };
	}
}
