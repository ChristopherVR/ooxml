import { XMLParser } from 'fast-xml-parser';
import { describe, it, expect } from 'vitest';

import type { XmlObject } from '../../types';
import { PptxTableDataParser } from './PptxTableDataParser';
import { extractTableCellParagraphs } from './table-cell-paragraphs';

const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: '@_',
	parseAttributeValue: false,
	parseTagValue: false,
	trimValues: false,
});

function ensureArray(value: unknown): unknown[] {
	if (value === undefined || value === null) {
		return [];
	}
	return Array.isArray(value) ? value : [value];
}

function parseCell(xml: string): XmlObject {
	return (parser.parse(xml) as XmlObject)['a:tc'] as XmlObject;
}

describe('extractTableCellParagraphs', () => {
	it("reads each paragraph's own alignment, indent and spacing", () => {
		const cell = parseCell(
			'<a:tc><a:txBody><a:bodyPr/>' +
				'<a:p><a:pPr algn="r" marL="95250" indent="-47625">' +
				'<a:lnSpc><a:spcPts val="960"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef>' +
				'</a:pPr><a:r><a:rPr sz="800"/><a:t>one</a:t></a:r></a:p>' +
				'<a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>' +
				'<a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr>' +
				'<a:r><a:rPr sz="1200"/><a:t>two</a:t></a:r></a:p>' +
				'<a:p><a:r><a:t>three</a:t></a:r></a:p>' +
				'</a:txBody></a:tc>',
		);
		expect(extractTableCellParagraphs(cell, { ensureArray })).toStrictEqual([
			{
				align: 'right',
				paragraphMarginLeft: 10,
				paragraphIndent: -5,
				lineSpacingExactPt: 9.6,
				paragraphSpacingBefore: 8,
			},
			// 50% of one 12pt (16px) line, which is 1.2x the font size.
			{ align: 'center', lineSpacing: 1.5, paragraphSpacingAfter: 0.5 * 16 * 1.2 },
			{},
		]);
	});

	it('sizes percentage spacing from the end properties, then the table default', () => {
		const cell = parseCell(
			'<a:tc><a:txBody><a:bodyPr/>' +
				'<a:p><a:pPr><a:spcAft><a:spcPct val="100000"/></a:spcAft></a:pPr>' +
				'<a:endParaRPr sz="900"/></a:p>' +
				'<a:p><a:pPr><a:spcAft><a:spcPct val="100000"/></a:spcAft></a:pPr>' +
				'<a:r><a:t>x</a:t></a:r></a:p>' +
				'</a:txBody></a:tc>',
		);
		expect(extractTableCellParagraphs(cell, { ensureArray }, 18)).toStrictEqual([
			{ endParaFontSize: 9, paragraphSpacingAfter: 12 * 1.2 },
			{ paragraphSpacingAfter: 24 * 1.2 },
		]);
		// With no size anywhere a percentage cannot be resolved.
		expect(extractTableCellParagraphs(cell, { ensureArray })).toStrictEqual([
			{ endParaFontSize: 9, paragraphSpacingAfter: 12 * 1.2 },
			{},
		]);
	});

	it('reads a right-to-left paragraph and an empty line size', () => {
		const cell = parseCell(
			'<a:tc><a:txBody><a:bodyPr/>' +
				'<a:p><a:pPr rtl="1"/><a:r><a:t>x</a:t></a:r></a:p>' +
				'<a:p><a:endParaRPr sz="600"/></a:p>' +
				'<a:p><a:br/><a:endParaRPr sz="600"/></a:p>' +
				'</a:txBody></a:tc>',
		);
		expect(extractTableCellParagraphs(cell, { ensureArray })).toStrictEqual([
			{ rtl: true },
			{ endParaFontSize: 6 },
			// A paragraph holding a line break is not empty.
			{},
		]);
	});

	it('returns undefined when no paragraph formats itself', () => {
		const cell = parseCell(
			'<a:tc><a:txBody><a:bodyPr/><a:p><a:pPr lvl="0"/><a:r><a:t>x</a:t></a:r></a:p>' +
				'<a:p><a:r><a:t>y</a:t></a:r></a:p></a:txBody></a:tc>',
		);
		expect(extractTableCellParagraphs(cell, { ensureArray })).toBeUndefined();
		expect(extractTableCellParagraphs(undefined, { ensureArray })).toBeUndefined();
	});
});

describe('PptxTableDataParser paragraphs', () => {
	it('attaches paragraphs to a cell with runs and leaves plain cells as they were', () => {
		const table = new PptxTableDataParser({
			emuPerPx: 9525,
			ensureArray,
			parseColor: () => undefined,
		}).parseTableData(
			parser.parse(
				'<a:tbl><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/><a:gridCol w="100"/>' +
					'</a:tblGrid><a:tr h="100">' +
					'<a:tc><a:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/><a:r><a:t>a</a:t></a:r></a:p>' +
					'</a:txBody></a:tc>' +
					'<a:tc><a:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/></a:p></a:txBody></a:tc>' +
					'<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>b</a:t></a:r></a:p></a:txBody></a:tc>' +
					'</a:tr></a:tbl>',
			) as XmlObject,
		);
		const [formatted, empty, plain] = table?.rows[0].cells ?? [];
		expect(formatted.paragraphs).toStrictEqual([{ align: 'center' }]);
		// The renderer draws paragraphs only alongside runs, so a cell without
		// runs does not carry them.
		expect(empty).not.toHaveProperty('paragraphs');
		expect(plain).not.toHaveProperty('paragraphs');
	});
});
