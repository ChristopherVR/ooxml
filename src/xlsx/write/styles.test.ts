import { describe, expect, it } from 'vitest';
import { DEFAULT_STYLES, miniPackage, roundTrip, ws } from './mini-package-fixtures.js';

// The review's t1: A1 typed as '123 in Excel (text format, quotePrefix), B1 a pivot button.
const STYLES = DEFAULT_STYLES.replace(
	'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
	'<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" pivotButton="1"/></cellXfs>',
);
const SHEET = ws(
	'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>123</t></is></c><c r="B1" s="2" t="inlineStr"><is><t>Region</t></is></c></row>',
);

describe('cell format flags', () => {
	it('reads and writes quotePrefix and pivotButton', async () => {
		const { wb, again, part } = await roundTrip(
			await miniPackage({ sheets: [{ name: 'S', xml: SHEET }], styles: STYLES }),
		);
		const styleOf = (book: typeof wb, col: number) =>
			book.styles[book.sheets[0]?.rows.get(0)?.get(col)?.styleId ?? 0];
		expect(styleOf(wb, 0)).toMatchObject({ quotePrefix: true, numFmt: '@' });
		expect(styleOf(wb, 1)?.pivotButton).toBe(true);
		expect(styleOf(again, 0)?.quotePrefix).toBe(true);
		expect(styleOf(again, 1)?.pivotButton).toBe(true);
		const xml = (await part('xl/styles.xml')) ?? '';
		expect(xml).toMatch(/<xf numFmtId="49"[^>]*quotePrefix="1"/);
		expect(xml).toMatch(/<xf [^>]*pivotButton="1"/);
	});
});
