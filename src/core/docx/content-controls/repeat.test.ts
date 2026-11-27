// Repeated content controls: Word's repeating sections at block and row level, plain controls and
// table rows named in `repeat`, sub-rows, and an empty list; saved and read back.
import { describe, expect, it } from 'vitest';
import { documentParagraphs } from '../cell-blocks';
import type { Block, Table } from '../model';
import { fillContentControls, readContentControls } from './index';
import {
	cell,
	item,
	paragraph,
	run,
	saveAndReload,
	section,
	sdt,
	styled,
	table,
	template,
} from './test-fixture';

const texts = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((entry) => entry.runs.map((r) => r.text).join(''));
const tableOf = (blocks: readonly Block[]) =>
	blocks.find((block): block is Table => block.type === 'table')!;
const rowTexts = (blocks: readonly Block[]) =>
	tableOf(blocks).rows.map((row) => row.map((c) => texts(c.paragraphs).join('|')).join(' / '));
const SHADED =
	'<w:shd w:val="clear" w:fill="FFEE00"/><w:tcBorders><w:top w:val="double"/></w:tcBorders>';
const ITEMS = [
	{ item_template_desc: 'Pen', qty: '2' },
	{ item_template_desc: 'Ink', qty: '5' },
	{ item_template_desc: 'Pad', qty: '1' },
];

describe('repeating content controls', () => {
	it('repeats a block repeating section once per item and retags the copies', async () => {
		const loaded = await template(
			paragraph(run('Before')) +
				sdt(
					'items',
					sdt(
						undefined,
						paragraph(
							sdt('item_template_desc', run('d')) + run(' x ') + sdt('item_template_qty', run('q')),
						),
						item,
					),
					section,
				) +
				paragraph(run('After')),
		);
		const filled = fillContentControls(loaded.model, { values: {}, repeat: { items: ITEMS } });
		expect(filled.results[0]).toMatchObject({ tag: 'items', status: 'repeated', items: 3 });
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks)).toEqual([
			'Before',
			'Pen x 2',
			'Ink x 5',
			'Pad x 1',
			'After',
		]);
		// One section holding three items.
		expect(reopened.xml.match(/<w15:repeatingSection\b/g)).toHaveLength(1);
		expect(reopened.xml.match(/<w15:repeatingSectionItem\b/g)).toHaveLength(3);
		const items = reopened.model.blocks.slice(1, 4);
		expect(items.map((block) => block.outerContentControls?.[0]?.tag)).toEqual([
			'items',
			'items',
			'items',
		]);
		expect(new Set(items.map((block) => block.contentControl?.id)).size).toBe(3);
		expect(readContentControls(reopened.model)).toMatchObject({
			item_1_desc: 'Pen',
			item_2_qty: '5',
			item_3_desc: 'Pad',
		});
		const ids = reopened.xml.match(/<w:id w:val="(\d+)"\/>/g)!;
		expect(new Set(ids).size).toBe(ids.length);
	});

	it('copies a plain block control per item, keeping its tag', async () => {
		const loaded = await template(
			sdt('line', paragraph(styled('', '<w:b/>') + sdt('line_template_name', run('n')))) +
				paragraph(run('end')),
		);
		const filled = fillContentControls(loaded.model, {
			values: {},
			repeat: { line: [{ name: 'A' }, { name: 'B' }] },
		});
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks)).toEqual(['A', 'B', 'end']);
		expect(reopened.model.blocks.slice(0, 2).map((block) => block.contentControl?.tag)).toEqual([
			'line',
			'line',
		]);
		expect(readContentControls(reopened.model)).toMatchObject({
			line_1_name: 'A',
			line_2_name: 'B',
		});
	});

	it('repeats a table row holding a control, with its formatting, and per-item sub-rows', async () => {
		const loaded = await template(
			table(
				`<w:tr>${cell(paragraph(run('Item')))}${cell(paragraph(run('Qty')))}</w:tr>` +
					`<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>${cell(paragraph(sdt('row_template_desc', run('d'))), SHADED)}${cell(paragraph(sdt('row_template_qty', run('q'))), SHADED)}</w:tr>` +
					`<w:tr>${cell(paragraph(run('Discount')))}${cell(paragraph(sdt('row_template_discount', run('x'))))}</w:tr>` +
					`<w:tr>${cell(paragraph(run('Total')))}${cell(paragraph(sdt('total', run('t'))))}</w:tr>`,
			),
		);
		const filled = fillContentControls(loaded.model, {
			values: { total: '8' },
			repeat: {
				row: [
					{ desc: 'Pen', qty: '2', discount: '10%' },
					{ desc: 'Ink', qty: '5' },
				],
			},
		});
		const reopened = await saveAndReload(loaded, filled.model);
		expect(rowTexts(reopened.model.blocks)).toEqual([
			'Item / Qty',
			'Pen / 2',
			'Discount / 10%',
			'Ink / 5',
			'Total / 8',
		]);
		expect(reopened.xml.match(/w:fill="FFEE00"/g)).toHaveLength(4);
		expect(reopened.xml.match(/<w:top w:val="double"\/>/g)).toHaveLength(4);
		expect(reopened.xml.match(/<w:trHeight w:val="400"\/>/g)).toHaveLength(2);
		expect(readContentControls(reopened.model)).toMatchObject({
			row_1_desc: 'Pen',
			row_1_discount: '10%',
			row_2_qty: '5',
		});
	});

	it('repeats the rows of a row-level repeating section inside the section', async () => {
		const loaded = await template(
			table(
				`<w:tr>${cell(paragraph(run('Head')))}${cell(paragraph(run('')))}</w:tr>` +
					sdt(
						'fees',
						sdt(
							undefined,
							`<w:tr>${cell(paragraph(sdt('fee_template_name', run('n'))), SHADED)}${cell(paragraph(sdt('fee_template_sum', run('s'))))}</w:tr>`,
							item,
						),
						section,
					),
			),
		);
		const filled = fillContentControls(loaded.model, {
			values: {},
			repeat: {
				fees: [
					{ name: 'A', sum: '1' },
					{ name: 'B', sum: '2' },
				],
			},
		});
		const reopened = await saveAndReload(loaded, filled.model);
		expect(rowTexts(reopened.model.blocks)).toEqual(['Head / ', 'A / 1', 'B / 2']);
		expect(reopened.xml.match(/<w15:repeatingSection\b/g)).toHaveLength(1);
		expect(reopened.xml.match(/<w15:repeatingSectionItem\b/g)).toHaveLength(2);
		const rows = tableOf(reopened.model.blocks).rowProperties!;
		expect(rows.slice(1).map((props) => props.outerContentControls?.[0]?.tag)).toEqual([
			'fees',
			'fees',
		]);
		expect(reopened.xml.match(/w:fill="FFEE00"/g)).toHaveLength(2);
	});

	it('removes a repeated control without items and leaves a valid table', async () => {
		const loaded = await template(
			table(
				`<w:tr>${cell(paragraph(run('Head')))}${cell(paragraph(run('')))}</w:tr>` +
					`<w:tr>${cell(paragraph(sdt('row', run('r'))))}${cell(paragraph(run('')))}</w:tr>`,
			) +
				sdt('block', paragraph(run('gone'))) +
				paragraph(run('end')),
		);
		const filled = fillContentControls(loaded.model, {
			values: {},
			repeat: { row: [], block: [] },
		});
		expect(filled.results.map((result) => [result.tag, result.status, result.items])).toEqual([
			['row', 'removed', 0],
			['block', 'removed', 0],
		]);
		const reopened = await saveAndReload(loaded, filled.model);
		expect(rowTexts(reopened.model.blocks)).toEqual(['Head / ']);
		expect(texts(reopened.model.blocks)).toEqual(['Head', '', 'end']);
		expect(reopened.xml).not.toContain('<w:sdt>');
	});
});
