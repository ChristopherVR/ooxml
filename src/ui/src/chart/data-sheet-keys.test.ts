import { chartDataSheetKeyAction } from './data-sheet-keys';

// A sheet with two series and two categories: 3 x 3 grid cells, the corner (0, 0) excluded.
const act = (
	key: string,
	mode: 'ready' | 'enter' | 'edit',
	row: number,
	col: number,
	shiftKey = false,
) => chartDataSheetKeyAction({ key, shiftKey }, mode, row, col, 3, 3);

describe('chart data sheet keys', () => {
	it('moves with arrows unless the cell is in F2 edit mode', () => {
		expect(act('ArrowDown', 'ready', 1, 1)).toEqual({ type: 'move', row: 2, col: 1 });
		expect(act('ArrowRight', 'enter', 1, 1)).toEqual({ type: 'move', row: 1, col: 2 });
		expect(act('ArrowLeft', 'edit', 1, 1)).toEqual({ type: 'native' });
		// The corner cell does not exist and the edges hold.
		expect(act('ArrowUp', 'ready', 1, 0)).toEqual({ type: 'move', row: 1, col: 0 });
		expect(act('ArrowDown', 'ready', 2, 2)).toEqual({ type: 'move', row: 2, col: 2 });
	});

	it('commits and moves with Enter and Tab, leaving the sheet past its edges', () => {
		expect(act('Enter', 'enter', 0, 1)).toEqual({ type: 'move', row: 1, col: 1 });
		expect(act('Enter', 'enter', 2, 1, true)).toEqual({ type: 'move', row: 1, col: 1 });
		expect(act('Tab', 'ready', 1, 0)).toEqual({ type: 'move', row: 1, col: 1 });
		expect(act('Tab', 'ready', 1, 2)).toEqual({ type: 'native' });
		expect(act('Tab', 'ready', 0, 1, true)).toEqual({ type: 'native' });
	});

	it('enters edit and typing modes and reverts with Escape', () => {
		expect(act('F2', 'ready', 1, 1)).toEqual({ type: 'edit' });
		expect(act('7', 'ready', 1, 1)).toEqual({ type: 'typing' });
		expect(act('7', 'enter', 1, 1)).toEqual({ type: 'native' });
		expect(act('Backspace', 'ready', 1, 1)).toEqual({ type: 'typing' });
		expect(act('Escape', 'enter', 1, 1)).toEqual({ type: 'revert' });
		expect(chartDataSheetKeyAction({ key: 'z', ctrlKey: true }, 'ready', 1, 1, 3, 3)).toEqual({
			type: 'native',
		});
	});
});
