import { disableAutoUnmount, enableAutoUnmount, mount } from '@vue/test-utils';
import type { PptxTableData } from 'pptx-viewer-core';
import { afterAll, afterEach, describe, expect, it } from 'vitest';

import TableCellFormattingPanel from './TableCellFormattingPanel.vue';

enableAutoUnmount(afterEach);
afterAll(disableAutoUnmount);

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
		const wrapper = mount(TableCellFormattingPanel, {
			props: { tableData, rowIndex: 0, columnIndex: 0, canEdit: true },
		});
		const values = wrapper
			.findAll('input[type="number"]')
			.map((input) => (input.element as HTMLInputElement).value);
		expect(values.filter((value) => value === '2.03')).toHaveLength(4);
		expect(values.filter((value) => value === '5.33')).toHaveLength(4);
		expect(values.filter((value) => /\.\d{3,}/u.test(value))).toStrictEqual([]);
	});
});
