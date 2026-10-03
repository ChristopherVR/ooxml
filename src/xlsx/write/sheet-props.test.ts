import { describe, expect, it } from 'vitest';
import { miniPackage, roundTrip, ws } from './mini-package-fixtures.js';

const ROWS =
	'<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>';
const sheet = (before: string, after: string) =>
	miniPackage({
		sheets: [
			{
				name: 'S',
				xml: ws(ROWS, before, after).replace(
					'<worksheet ',
					'<worksheet xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ',
				),
			},
		],
	});

const FILTER_COLUMNS = [
	'<filterColumn colId="0"><filters blank="1"><dateGroupItem year="2024" month="1" dateTimeGrouping="month"/></filters></filterColumn>',
	'<filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="1"/><customFilter operator="lessThan" val="9"/></customFilters></filterColumn>',
	'<filterColumn colId="2"><top10 top="0" percent="1" val="10"/></filterColumn>',
	'<filterColumn colId="3"><dynamicFilter type="aboveAverage"/></filterColumn>',
	'<filterColumn colId="4"><colorFilter dxfId="0"/></filterColumn>',
	'<filterColumn colId="5"><iconFilter iconSet="3Arrows" iconId="1"/></filterColumn>',
];
const AUTO_FILTER = `<autoFilter ref="A1:F3">${FILTER_COLUMNS.join('')}</autoFilter>`;

describe('autoFilter criteria', () => {
	it('keeps every kind of filterColumn on an unedited round trip', async () => {
		const { part } = await roundTrip(await sheet('', AUTO_FILTER));
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		for (const column of FILTER_COLUMNS) expect(xml).toContain(column);
	});

	it('re-indexes kept criteria when the range moves and never writes empty columns', async () => {
		const { part, again } = await roundTrip(await sheet('', AUTO_FILTER), (wb) => {
			const filter = wb.sheets[0]!.autoFilter!;
			// What a column insert before the filter does: the range moves, offsets stay.
			filter.range = { start: { row: 0, col: 1 }, end: { row: 2, col: 6 } };
			// Dropping the first column: the rest move one place left.
			filter.columns = filter.columns!.slice(1).map((c) => ({ ...c, offset: c.offset - 1 }));
			filter.range = { start: { row: 0, col: 2 }, end: { row: 2, col: 6 } };
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain('<autoFilter ref="C1:G3">');
		expect(xml).toContain(
			'<filterColumn colId="0"><customFilters and="1"><customFilter operator="greaterThan" val="1"/>',
		);
		expect(xml).toContain('<filterColumn colId="4"><iconFilter iconSet="3Arrows" iconId="1"/>');
		expect(xml).not.toMatch(/<filterColumn colId="\d+"\/>/);
		expect(again.sheets[0]?.autoFilter?.columns?.length).toBe(5);
	});

	it('writes an edited column from the model', async () => {
		const { part } = await roundTrip(await sheet('', AUTO_FILTER), (wb) => {
			const filter = wb.sheets[0]!.autoFilter!;
			filter.columns = [{ offset: 1, values: ['2'] }];
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain(
			'<autoFilter ref="A1:F3"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>',
		);
	});
});

describe('sheetFormatPr', () => {
	const FORMAT =
		'<sheetFormatPr baseColWidth="10" defaultRowHeight="15" zeroHeight="1" thickTop="1" thickBottom="1" x14ac:dyDescent="0.25"/>';

	it('keeps zeroHeight, baseColWidth, thick borders and x14ac:dyDescent when regenerated', async () => {
		const { part, again } = await roundTrip(await sheet(FORMAT, ''), (wb) => {
			wb.sheets[0]!.defaultColWidth = 12;
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		const format = /<sheetFormatPr[^>]*\/>/.exec(xml)?.[0] ?? '';
		for (const attr of [
			'baseColWidth="10"',
			'zeroHeight="1"',
			'thickTop="1"',
			'thickBottom="1"',
			'x14ac:dyDescent="0.25"',
			'defaultColWidth="12"',
		])
			expect(format).toContain(attr);
		expect(again.sheets[0]?.format).toEqual({
			zeroHeight: true,
			thickTop: true,
			thickBottom: true,
			baseColWidth: 10,
		});
	});
});

describe('headerFooter', () => {
	const HF =
		'<headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0"><oddHeader>&amp;CODD</oddHeader><oddFooter>&amp;Pf</oddFooter><evenHeader>&amp;CEVEN</evenHeader><evenFooter>ef</evenFooter><firstHeader>&amp;CFIRST</firstHeader><firstFooter>ff</firstFooter></headerFooter>';

	it('reads every header and keeps the rest when one is edited', async () => {
		const { wb, part } = await roundTrip(await sheet('', HF), (book) => {
			book.sheets[0]!.pageSetup!.header = '&CNEW';
		});
		expect(wb.sheets[0]?.pageSetup).toMatchObject({
			differentFirst: true,
			differentOddEven: true,
			firstHeader: '&CFIRST',
			evenHeader: '&CEVEN',
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain(
			'<headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0"><oddHeader>&amp;CNEW</oddHeader><oddFooter>&amp;Pf</oddFooter><evenHeader>&amp;CEVEN</evenHeader><evenFooter>ef</evenFooter><firstHeader>&amp;CFIRST</firstHeader><firstFooter>ff</firstFooter></headerFooter>',
		);
	});
});

describe('pageSetup', () => {
	const SETUP =
		'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" firstPageNumber="5" useFirstPageNumber="1" pageOrder="overThenDown" blackAndWhite="1" draft="1" cellComments="atEnd" errors="blank" horizontalDpi="600" verticalDpi="300" copies="2"/>';

	it('patches the orientation and keeps every other attribute', async () => {
		const { part } = await roundTrip(await sheet('', SETUP), (wb) => {
			wb.sheets[0]!.pageSetup!.orientation = 'landscape';
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		const setup = /<pageSetup[^>]*\/>/.exec(xml)?.[0] ?? '';
		for (const attr of [
			'paperSize="9"',
			'orientation="landscape"',
			'firstPageNumber="5"',
			'useFirstPageNumber="1"',
			'pageOrder="overThenDown"',
			'blackAndWhite="1"',
			'draft="1"',
			'cellComments="atEnd"',
			'errors="blank"',
			'horizontalDpi="600"',
			'verticalDpi="300"',
			'copies="2"',
		])
			expect(setup).toContain(attr);
	});
});
