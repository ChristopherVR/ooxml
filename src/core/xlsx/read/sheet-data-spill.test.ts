import { describe, expect, it } from 'vitest';
import { NS, elements, parseXml, type XmlElement } from '../../xml/index';
import { createWorksheet } from '../workbook';
import { readSheetData } from './sheet-data';

function read(data: string, dynamic = new Set([1])) {
	const xml = `<worksheet xmlns="${NS.x}"><sheetData>${data}</sheetData></worksheet>`;
	const sheetData = elements(parseXml(xml).documentElement).find(
		(node) => node.localName === 'sheetData',
	) as XmlElement;
	const sheet = createWorksheet('S', 1);
	readSheetData(
		{
			sharedStrings: [],
			xfMap: new Map(),
			date1904: false,
			palette: [],
			dynamicCells: dynamic,
			warn: () => {},
		} as never,
		sheetData,
		sheet,
	);
	return sheet;
}

describe('spill marking', () => {
	it('marks cached spill cells for a ref covering the whole sheet, quickly', () => {
		const started = Date.now();
		const sheet = read(
			'<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:XFD1048576">SEQUENCE(2)</f><v>1</v></c></row>' +
				'<row r="2"><c r="A2"><v>2</v></c><c r="C2"><v>3</v></c></row>',
		);
		expect(Date.now() - started).toBeLessThan(1000);
		expect(
			(sheet.rows.get(1)?.get(0) as { spillAnchor?: unknown } | undefined)?.spillAnchor,
		).toEqual({
			row: 0,
			col: 0,
		});
		expect(
			(sheet.rows.get(1)?.get(2) as { spillAnchor?: unknown } | undefined)?.spillAnchor,
		).toEqual({
			row: 0,
			col: 0,
		});
		expect(
			(sheet.rows.get(0)?.get(0) as { spillAnchor?: unknown } | undefined)?.spillAnchor,
		).toBeUndefined();
	});

	it('leaves cells outside the ref alone', () => {
		const sheet = read(
			'<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">SEQUENCE(2)</f><v>1</v></c></row>' +
				'<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>3</v></c></row>',
		);
		expect(
			(sheet.rows.get(1)?.get(0) as { spillAnchor?: unknown } | undefined)?.spillAnchor,
		).toBeDefined();
		expect(
			(sheet.rows.get(1)?.get(1) as { spillAnchor?: unknown } | undefined)?.spillAnchor,
		).toBeUndefined();
	});
});
