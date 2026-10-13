// @vitest-environment happy-dom
import type { PptxTableData } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { TableCellFormattingPanel } from './TableCellFormattingPanel';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

describe('tableCellFormattingPanel', () => {
	it('aligns every paragraph of the cell', () => {
		const tableData: PptxTableData = {
			columnWidths: [1],
			rows: [
				{
					cells: [
						{
							text: 'one\ntwo',
							style: { align: 'right' },
							paragraphs: [{ align: 'right' }, { align: 'center' }],
						},
					],
				},
			],
		};
		const onUpdateTableData = vi.fn<(patch: Partial<PptxTableData>) => void>();
		act(() => {
			root.render(
				<TableCellFormattingPanel
					tableData={tableData}
					tableEditorState={{ rowIndex: 0, columnIndex: 0, elementId: 'table' }}
					canEdit
					onUpdateTableData={onUpdateTableData}
					onUpdateMergeRows={() => {}}
				/>,
			);
		});
		const left = [...container.querySelectorAll('button')].find(
			(button) => button.textContent === 'L',
		);
		act(() => left!.click());
		expect(onUpdateTableData.mock.calls[0][0].rows?.[0].cells[0]).toMatchObject({
			style: { align: 'left' },
			paragraphs: [{ align: 'left' }, { align: 'left' }],
		});
	});
	it('shows EMU-derived widths and margins with two decimals at most', () => {
		// 1.52pt borders and 4pt margins, read back from EMU as px.
		const border = 19304 / 9525;
		const margin = 50800 / 9525;
		const tableData: PptxTableData = {
			columnWidths: [1],
			rows: [
				{
					cells: [
						{
							text: 'one',
							style: {
								borderTopWidth: border,
								borderBottomWidth: border,
								borderLeftWidth: border,
								borderRightWidth: border,
								marginTop: margin,
								marginBottom: margin,
								marginLeft: margin,
								marginRight: margin,
							},
						},
					],
				},
			],
		};
		act(() => {
			root.render(
				<TableCellFormattingPanel
					tableData={tableData}
					tableEditorState={{ rowIndex: 0, columnIndex: 0, elementId: 'table' }}
					canEdit
					onUpdateTableData={() => {}}
					onUpdateMergeRows={() => {}}
				/>,
			);
		});
		const values = [...container.querySelectorAll<HTMLInputElement>('input[type="number"]')].map(
			(input) => input.value,
		);
		expect(values.filter((value) => value === '2.03')).toHaveLength(4);
		expect(values.filter((value) => value === '5.33')).toHaveLength(4);
		expect(values.filter((value) => /\.\d{3,}/u.test(value))).toStrictEqual([]);
	});
});
