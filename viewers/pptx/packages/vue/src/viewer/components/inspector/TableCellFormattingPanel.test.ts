import { mount } from '@vue/test-utils';
import type { PptxTableData } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import TableCellFormattingPanel from './TableCellFormattingPanel.vue';

describe('tableCellFormattingPanel', () => {
	it('aligns every paragraph of the cell', async () => {
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
		const wrapper = mount(TableCellFormattingPanel, {
			props: { tableData, rowIndex: 0, columnIndex: 0, canEdit: true },
		});
		const left = wrapper.findAll('button').find((button) => button.text() === 'L');
		await left!.trigger('click');
		const [patch] = wrapper.emitted('update')![0] as [Partial<PptxTableData>];
		expect(patch.rows?.[0].cells[0]).toMatchObject({
			style: { align: 'left' },
			paragraphs: [{ align: 'left' }, { align: 'left' }],
		});
	});
});
