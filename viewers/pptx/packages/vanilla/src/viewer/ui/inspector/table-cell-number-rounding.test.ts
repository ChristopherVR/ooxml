import { describe, expect, it, vi } from 'vitest';

import { createTableCellFillControls } from './table-cell-fill-controls';
import { createTableCellFormatting } from './table-cell-formatting';
import { createTableStructureControls } from './table-structure-controls';
import type { InspectorHandlers, InspectorState } from './types';

// 1.52pt borders and 4pt margins, read back from EMU as px.
const BORDER_PX = 19304 / 9525;
const MARGIN_PX = 50800 / 9525;

const STATE = {
	isTable: true,
	selectedTableCell: { row: 0, column: 0 },
	selectedTableCells: [],
	tableColumnWidths: [1],
	tableRowHeights: [370840 / 9525],
	tableCellStyle: {
		fontSize: 10.5,
		borderTopWidth: BORDER_PX,
		marginLeft: MARGIN_PX,
		gradientFillAngle: 33.33333333,
	},
} as unknown as InspectorState;

const HANDLERS = {} as InspectorHandlers;
const identity = (key: string): string => key;

function numberInputs(root: HTMLElement): string[] {
	return Array.from(root.querySelectorAll<HTMLInputElement>('input[type="number"]')).map(
		(input) => input.value,
	);
}

describe('table cell inspector number display', () => {
	it('shows border width and padding with two decimals at most', () => {
		const panel = createTableCellFormatting(document, identity, HANDLERS);
		panel.update(STATE);

		expect(numberInputs(panel.el)).toStrictEqual(['10.5', '2.03', '5.33']);
	});

	it('commits a typed value as typed, without rounding it', () => {
		const setTableCellStyles = vi.fn();
		const panel = createTableCellFormatting(document, identity, {
			setTableCellStyles,
		} as unknown as InspectorHandlers);
		panel.update(STATE);
		const borderWidth = panel.el.querySelectorAll<HTMLInputElement>('input[type="number"]')[1]!;
		borderWidth.value = '2.025';
		borderWidth.dispatchEvent(new Event('change'));

		expect(setTableCellStyles).toHaveBeenCalledWith(
			[{ row: 0, column: 0 }],
			expect.objectContaining({ borderTopWidth: 2.025 }),
		);
	});

	it('rounds the fill panel angle and edge width', () => {
		const panel = createTableCellFillControls(document, identity, HANDLERS);
		panel.update(STATE);

		expect(numberInputs(panel.el)).toStrictEqual(['33.33', '2.03']);
	});

	it('rounds the row height', () => {
		const panel = createTableStructureControls(document, identity, HANDLERS);
		panel.update(STATE);

		expect(numberInputs(panel.el)).toStrictEqual(['100', '38.93']);
	});
});
