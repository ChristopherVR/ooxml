import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxElement, PptxTableData, XmlObject } from '../../types';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';
import { tableStyleIdChanged } from './table-style-change-cell-clear';

const OLD_STYLE = '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}';
const NEW_STYLE = '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}';

class TableSaveRuntime extends PptxHandlerRuntime {
	public save(shape: XmlObject, data: PptxTableData): void {
		this.serializeTableDataToXml(shape, data);
	}
}

const line = (hex: string): XmlObject => ({
	'@_w': '12700',
	'a:solidFill': { 'a:srgbClr': { '@_val': hex } },
});

/** A one-row, two-cell table: the first cell has a fill and four borders. */
function tableShape(styleId: string): XmlObject {
	const cell = (tcPr: XmlObject): XmlObject => ({
		'a:txBody': { 'a:bodyPr': {}, 'a:p': { 'a:r': { 'a:rPr': { '@_b': '1' }, 'a:t': 'x' } } },
		'a:tcPr': tcPr,
	});
	return {
		'a:graphic': {
			'a:graphicData': {
				'a:tbl': {
					'a:tblPr': { '@_firstRow': '1', 'a:tableStyleId': styleId },
					'a:tblGrid': { 'a:gridCol': [{ '@_w': '100' }, { '@_w': '100' }] },
					'a:tr': {
						'@_h': '100',
						'a:tc': [
							cell({
								'a:lnL': line('FF0000'),
								'a:lnR': line('FF0000'),
								'a:lnT': line('FF0000'),
								'a:lnB': { 'a:noFill': '' },
								'a:solidFill': { 'a:srgbClr': { '@_val': '00FF00' } },
							}),
							cell({ 'a:lnT': line('0000FF') }),
						],
					},
				},
			},
		},
	};
}

function tcPrs(shape: XmlObject): XmlObject[] {
	const tbl = (shape['a:graphic'] as XmlObject)['a:graphicData'] as XmlObject;
	const row = (tbl['a:tbl'] as XmlObject)['a:tr'] as XmlObject;
	return (row['a:tc'] as XmlObject[]).map((tc) => tc['a:tcPr'] as XmlObject);
}

/** Table data as a style pick leaves it: no direct fill or border, bold kept. */
function pickedData(styleId: string): PptxTableData {
	return {
		columnWidths: [0.5, 0.5],
		firstRowHeader: true,
		tableStyleId: styleId,
		rows: [{ height: 0, cells: [{ style: { bold: true } }, {}] }],
	};
}

describe('saving a table whose style changed', () => {
	it('drops the direct fills and borders the model no longer carries', () => {
		const shape = tableShape(OLD_STYLE);
		new TableSaveRuntime().save(shape, pickedData(NEW_STYLE));
		const [first, second] = tcPrs(shape);
		expect(Object.keys(first).filter((key) => key.startsWith('a:'))).toStrictEqual([]);
		expect(second['a:lnT']).toBeUndefined();
	});

	it('still writes a fill or border set after the pick', () => {
		const shape = tableShape(OLD_STYLE);
		const data = pickedData(NEW_STYLE);
		data.rows[0].cells[0].style = { backgroundColor: '#123456', borderTopWidth: 2 };
		new TableSaveRuntime().save(shape, data);
		const [first] = tcPrs(shape);
		expect(first['a:solidFill']).toStrictEqual({ 'a:srgbClr': { '@_val': '123456' } });
		expect(first['a:lnT']).toMatchObject({ '@_w': '19050' });
		expect(first['a:lnL']).toBeUndefined();
	});

	it('leaves the cells alone when the style is unchanged', () => {
		const shape = tableShape(OLD_STYLE);
		new TableSaveRuntime().save(shape, pickedData(OLD_STYLE.toLowerCase().slice(1, -1)));
		const [first, second] = tcPrs(shape);
		expect(first['a:lnL']).toBeDefined();
		expect(first['a:lnB']).toStrictEqual({ 'a:noFill': '' });
		expect(first['a:solidFill']).toBeDefined();
		expect(second['a:lnT']).toBeDefined();
	});

	it('clears with the same id when the data carries clearDirectCellFormatting', () => {
		const shape = tableShape(OLD_STYLE);
		new TableSaveRuntime().save(shape, {
			...pickedData(OLD_STYLE),
			clearDirectCellFormatting: true,
		});
		const [first, second] = tcPrs(shape);
		expect(first['a:solidFill']).toBeUndefined();
		expect(first['a:lnB']).toBeUndefined();
		expect(second['a:lnT']).toBeUndefined();
	});

	it('clears an explicit style removal (empty id) too', () => {
		const shape = tableShape(OLD_STYLE);
		new TableSaveRuntime().save(shape, pickedData(''));
		expect(tcPrs(shape)[0]['a:lnL']).toBeUndefined();
	});

	it('clears the cells reused by a structural rebuild', () => {
		const shape = tableShape(OLD_STYLE);
		const data = pickedData(NEW_STYLE);
		data.rows.push({ height: 0, cells: [{}, {}] });
		new TableSaveRuntime().save(shape, data);
		const tbl = ((shape['a:graphic'] as XmlObject)['a:graphicData'] as XmlObject)[
			'a:tbl'
		] as XmlObject;
		const rows = tbl['a:tr'] as XmlObject[];
		expect(rows).toHaveLength(2);
		const first = (rows[0]['a:tc'] as XmlObject[])[0]['a:tcPr'] as XmlObject;
		expect(first['a:lnL']).toBeUndefined();
		expect(first['a:solidFill']).toBeUndefined();
	});

	it('compares ids case- and brace-insensitively and keeps an undefined id', () => {
		const tblPr = { 'a:tableStyleId': OLD_STYLE };
		expect(tableStyleIdChanged(tblPr, OLD_STYLE.toLowerCase())).toBeFalsy();
		expect(tableStyleIdChanged(tblPr, undefined)).toBeFalsy();
		expect(tableStyleIdChanged(tblPr, NEW_STYLE)).toBeTruthy();
		expect(tableStyleIdChanged({}, NEW_STYLE)).toBeTruthy();
		expect(tableStyleIdChanged('', '')).toBeFalsy();
		// The legacy forms the parser also reads are the same style.
		expect(tableStyleIdChanged({ 'a:tblStyle': { '@_val': OLD_STYLE } }, OLD_STYLE)).toBeFalsy();
		expect(tableStyleIdChanged({ '@_tblStyle': OLD_STYLE }, OLD_STYLE)).toBeFalsy();
	});
});

const FIXTURE = fileURLToPath(
	new URL('../../../__tests__/fixtures/e2e/table-styling.pptx', import.meta.url),
);

/** Load a deck and return its handler, slides and first table. */
async function open(bytes: Uint8Array) {
	const handler = new PptxHandler();
	const deck = await handler.load(bytes.slice().buffer as ArrayBuffer);
	const table = deck.slides[0].elements.find((e) => e.type === 'table');
	if (!table || table.type !== 'table' || !table.tableData) {
		throw new Error('expected a table');
	}
	return { handler, slides: deck.slides, table, tableData: table.tableData };
}

/** Save `opened` with its table's data replaced by `tableData`. */
function saveWith(opened: Awaited<ReturnType<typeof open>>, tableData: PptxTableData) {
	const [first, ...rest] = opened.slides;
	const elements = first.elements.map((e) =>
		e.id === opened.table.id ? ({ ...e, tableData } as PptxElement) : e,
	);
	return opened.handler.save([{ ...first, elements }, ...rest]);
}

describe('table style change round trip', () => {
	it('does not bring the old borders and fill back after reload', async () => {
		const deck = await open(new Uint8Array(readFileSync(FIXTURE)));
		// Give the first cell a direct fill and a border, and save.
		const formatted: PptxTableData = structuredClone(deck.tableData);
		formatted.rows[0].cells[0].style = {
			...formatted.rows[0].cells[0].style,
			backgroundColor: '#00FF00',
			borderTopWidth: 3,
			borderTopColor: '#FF0000',
		};
		const reopened = await open(await saveWith(deck, formatted));
		expect(reopened.tableData.rows[0].cells[0].style?.borderTopColor).toBe('#FF0000');

		// A style pick: a new id, with the fill and border gone from the model.
		const picked: PptxTableData = structuredClone(reopened.tableData);
		picked.tableStyleId = NEW_STYLE;
		const style = { ...picked.rows[0].cells[0].style };
		for (const key of Object.keys(style).filter((k) => /^(background|fill|border)/u.test(k))) {
			delete (style as Record<string, unknown>)[key];
		}
		picked.rows[0].cells[0].style = style;
		const bytes = await saveWith(reopened, picked);

		const xml = await (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
		expect(xml).toContain(`<a:tableStyleId>${NEW_STYLE}</a:tableStyleId>`);
		expect(xml).not.toContain('FF0000');
		expect(xml).not.toContain('00FF00');
		const after = (await open(bytes)).tableData.rows[0].cells[0].style;
		expect(after?.backgroundColor).toBeUndefined();
		expect(after?.borderTopWidth).toBeUndefined();
	});
});
