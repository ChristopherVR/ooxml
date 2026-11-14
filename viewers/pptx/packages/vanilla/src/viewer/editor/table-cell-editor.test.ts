import type { TableCellEditOutcome, TableCellEditEnd } from 'ooxml-ui/pptx';
import { resolveTableCellEditEnd } from 'ooxml-ui/pptx';
import type { TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import type { EditorOps } from './editor-operations';
import { openTableCellEditor } from './table-cell-editor';

const element = (): TablePptxElement =>
	({
		id: 'table-1',
		type: 'table',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		tableData: {
			rows: [
				{
					cells: [{ text: 'Starter', textRuns: [{ text: 'Starter', bold: true }] }, { text: 'B' }],
				},
			],
			columnWidths: [0.5, 0.5],
		},
	}) as TablePptxElement;

function setup() {
	const host = document.createElement('div');
	host.dataset.elementId = 'table-1';
	const cells = ['Starter', 'B'].map((text, index) => {
		const cell = document.createElement('td');
		cell.dataset.rowIndex = '0';
		cell.dataset.cellIndex = String(index);
		cell.textContent = text;
		host.appendChild(cell);
		return cell;
	});
	document.body.appendChild(host);
	const endTableCellEdit = vi.fn(
		(_id: string, row: number, column: number, end: TableCellEditEnd): TableCellEditOutcome =>
			resolveTableCellEditEnd(element(), row, column, end),
	);
	const session = openTableCellEditor({
		doc: document,
		cell: cells[0]!,
		element: element(),
		row: 0,
		column: 0,
		ops: { endTableCellEdit, selectedElement: () => element() } as unknown as EditorOps,
	});
	const editor = () => document.querySelector<HTMLElement>('[data-table-cell-editor]')!;
	return { host, cells, endTableCellEdit, session, editor };
}

describe('openTableCellEditor (rich)', () => {
	it('mounts the shared rich editor with the cell runs and keeps their format', () => {
		const { host, cells, editor } = setup();
		expect(cells[0]!.contains(editor())).toBe(true);
		expect(editor().querySelector('span')!.style.fontWeight).toBe('bold');
		expect(editor().textContent).toBe('Starter');
		host.remove();
	});

	it('commits the edited runs and moves to the next cell on Tab', async () => {
		const { host, cells, endTableCellEdit, editor } = setup();
		editor().querySelector('span')!.textContent = 'Started';
		editor().dispatchEvent(new Event('input', { bubbles: true }));
		editor().dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
		const end = endTableCellEdit.mock.calls[0]![3];
		expect(end).toMatchObject({ reason: 'navigate', key: 'next' });
		expect(end.segments![0]).toMatchObject({ text: 'Started', style: { bold: true } });
		await Promise.resolve();
		expect(cells[1]!.querySelector('[data-table-cell-editor]')).not.toBeNull();
		host.remove();
	});

	it('closes without committing when cancelled, restoring the cell', () => {
		const { host, cells, endTableCellEdit, session } = setup();
		session.close(false);
		expect(endTableCellEdit).not.toHaveBeenCalled();
		expect(cells[0]!.textContent).toBe('Starter');
		host.remove();
	});
});
