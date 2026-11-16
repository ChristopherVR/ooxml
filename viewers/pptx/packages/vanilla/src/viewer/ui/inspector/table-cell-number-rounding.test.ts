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
	it('shows the border width and margin in points, rounded', () => {
		const panel = createTableCellFormatting(document, identity, HANDLERS);
		panel.update(STATE);

		const value = (id: string) =>
			panel.el.querySelector<HTMLInputElement>(`[data-testid="${id}"]`)!.value;
		expect(value('table-cell-fontSize')).toBe('10.5');
		expect(value('table-cell-borderTop-width')).toBe('1.52');
		expect(value('table-cell-marginLeft')).toBe('4');
	});

	it('commits a typed value as typed, without rounding it', () => {
		const setTableCellStyles = vi.fn();
		const panel = createTableCellFormatting(document, identity, {
			setTableCellStyles,
		} as unknown as InspectorHandlers);
		panel.update(STATE);
		const borderWidth = panel.el.querySelector<HTMLInputElement>(
			'[data-testid="table-cell-borderTop-width"]',
		)!;
		borderWidth.value = '1.5';
		borderWidth.dispatchEvent(new Event('change'));

		expect(setTableCellStyles).toHaveBeenCalledWith(
			[{ row: 0, column: 0 }],
			expect.objectContaining({ borderTopWidth: 2 }),
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
