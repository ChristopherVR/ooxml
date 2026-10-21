import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, elements, parseXml, type XmlElement } from '../../xml/index';
import { createWorksheet } from '../workbook';
import { parseDynamicArrayMetadata } from './metadata';
import { readZipParts, SourceIndex } from './package';
import { readSheetData, type CellContext } from './sheet-data';
import { splitSheetData } from './sheet-data-source';
import { parseSharedStrings } from './shared-strings';
import { parseStyles } from './styles';
import { parseWorkbookPart } from './workbook-part';

const FIXTURES = [
	'excel-features.xlsx',
	'excel-1904.xlsx',
	'excel-smartart.xlsx',
	'excel-sparklines.xlsx',
	'openpyxl-features.xlsx',
	'openpyxl-styles.xlsx',
	'openpyxl-1904.xlsx',
];

const context = (warnings: string[], source?: SourceIndex): CellContext => {
	const part = source?.workbookPart();
	const of = (type: string) => (part ? source?.targetOfType(part, type) : undefined);
	const text = (name: string | undefined) => (name ? source?.text(name) : undefined);
	const styles = parseStyles(text(of(RELATIONSHIP_TYPES.styles)), []);
	return {
		sharedStrings: parseSharedStrings(text(of(RELATIONSHIP_TYPES.sharedStrings)), styles.palette),
		xfMap: styles.xfMap,
		date1904: part ? parseWorkbookPart(source?.text(part) ?? '').date1904 : false,
		palette: styles.palette,
		dynamicCells: parseDynamicArrayMetadata(text(of(RELATIONSHIP_TYPES.sheetMetadata))),
		warn: (message) => warnings.push(message),
	};
};

/** Reads the cell data of a worksheet part through the DOM and through the split. */
function readBothWays(xml: string, source?: SourceIndex) {
	const split = splitSheetData(xml);
	const domData = elements(parseXml(xml).documentElement).find(
		(node) => node.localName === 'sheetData',
	) as XmlElement;
	const read = (data: XmlElement) => {
		const warnings: string[] = [];
		const sheet = createWorksheet('S', 1);
		readSheetData(context(warnings, source), data, sheet);
		return { rows: sheet.rows, rowInfo: sheet.rowInfo, warnings };
	};
	return { split, dom: read(domData), lite: split ? read(split.sheetData) : undefined };
}

describe('splitSheetData', () => {
	it.each(FIXTURES)('reads the cells of %s exactly as the DOM does', async (name) => {
		const bytes = readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name));
		const parts = await readZipParts(new Uint8Array(bytes));
		const source = new SourceIndex(parts);
		const sheets = [...parts.keys()].filter((part) => /^xl\/worksheets\/[^/]+\.xml$/.test(part));
		expect(sheets.length).toBeGreaterThan(0);
		for (const part of sheets) {
			const { split, dom, lite } = readBothWays(source.text(part) ?? '', source);
			expect(split, part).toBeDefined();
			expect(lite).toEqual(dom);
			// The rest of the part still parses, with an empty sheetData.
			const rest = parseXml(split?.xml ?? '').documentElement;
			expect(elements(rest).map((node) => node.localName)).toContain('sheetData');
		}
	});

	const sheet = (root: string, data: string, prefix = '') =>
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<${prefix}worksheet ${root}><${prefix}sheetData>${data}</${prefix}sheetData><${prefix}pageMargins left="0.7"/></${prefix}worksheet>`;
	const CELLS = [
		'<row r="1" spans="1:4" ht="20" customHeight="1" x14ac:dyDescent="0.25">',
		'<c r="A1" t="inlineStr"><is><t xml:space="preserve"> a &amp; b\r\n</t></is></c>',
		'<c r="B1" t="inlineStr"><is><r><rPr><b/><sz val="12"/></rPr><t>Bold</t></r><r><t>_x000D_x</t></r></is></c>',
		'<c r="C1" t="str"><f>"a"&amp;"b"</f><v>ab_x000A_</v></c>',
		'<c r="D1" t="e"><v>#N/A</v></c></row>',
		'<row r="3"><c r="A3"><f t="shared" ref="A3:A5" si="0">B3*2</f><v>2</v></c><c r="B3" s="1"><v>1</v></c></row>',
		'<row r="4"><c r="A4"><f t="shared" si="0"/><v>4</v></c><c r="B4" t="b"><v>1</v></c></row>',
		'<row r="5"><c r="A5"><f t="array" ref="A5:A6">B3:B4</f><v>1</v></c></row>',
	].join('');

	it('splits an unprefixed sheet in the SpreadsheetML default namespace', () => {
		const xml = sheet(`xmlns="${NS.x}" xmlns:x14ac="urn:x14ac"`, CELLS);
		const { split, dom, lite } = readBothWays(xml);
		expect(split).toBeDefined();
		expect(lite).toEqual(dom);
		expect(dom.rows.get(0)?.get(0)?.value).toBe(' a & b\n');
		expect(split?.xml).toContain('<sheetData></sheetData><pageMargins');
	});

	it('splits a sheet whose elements carry a prefix (Open XML SDK output)', () => {
		const xml = sheet(
			`xmlns:x="${NS.x}" xmlns:x14ac="urn:x14ac"`,
			CELLS.replace(/<(\/?)/g, '<$1x:'),
			'x:',
		);
		const { split, dom, lite } = readBothWays(xml);
		expect(split).toBeDefined();
		expect(lite).toEqual(dom);
		expect(dom.rows.get(0)?.get(1)?.richText?.[0]?.font?.bold).toBe(true);
	});

	it('reads a sheet in another namespace the same way (its cells keep no values)', () => {
		const strict = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
		const { split, dom, lite } = readBothWays(
			sheet(`xmlns="${strict}" xmlns:x14ac="urn:x14ac"`, CELLS),
		);
		expect(split).toBeDefined();
		expect(lite).toEqual(dom);
	});

	const COMMENTED: [string, (xml: string) => string][] = [
		[
			'a comment before the root',
			(xml) => xml.replace('\r\n<worksheet', '\r\n<!-- written by a tool -->\r\n<worksheet'),
		],
		[
			'a comment in the header',
			(xml) => xml.replace('<sheetData>', '<!-- <sheetData> --><sheetData>'),
		],
		[
			'a comment after the cell data',
			(xml) => xml.replace('<pageMargins', '<!-- </sheetData> --><pageMargins'),
		],
		['a comment between rows', (xml) => xml.replace('<row r="3">', '<!-- c --><row r="3">')],
		['a comment inside a row', (xml) => xml.replace('<c r="B3"', '<!-- in a row --><c r="B3"')],
		[
			'a comment inside a cell',
			(xml) => xml.replace('<c r="B4" t="b"><v>', '<c r="B4" t="b"><!-- x --><v>'),
		],
		['a comment inside a value', (xml) => xml.replace('<v>#N/A</v>', '<v>#N<!-- x -->/A</v>')],
		[
			'a comment holding markup',
			(xml) =>
				xml.replace(
					'<row r="4">',
					'<!-- <row r="9"><c r="A9"><v>9</v></c></row></sheetData> --><row r="4">',
				),
		],
		[
			'CDATA in inline text',
			(xml) => xml.replace('<t>Bold</t>', '<t><![CDATA[Bo<l>d & </sheetData>]]></t>'),
		],
		[
			'CDATA in a formula',
			(xml) => xml.replace('<f>"a"&amp;"b"</f>', '<f><![CDATA["a"&"b"]]></f>'),
		],
		[
			'CDATA next to text',
			(xml) => xml.replace('<v>ab_x000A_</v>', '<v>a<![CDATA[b]]>_x000A_</v>'),
		],
		['empty CDATA', (xml) => xml.replace('<v>4</v>', '<v>4<![CDATA[]]></v>')],
	];

	it.each(COMMENTED)('splits a sheet with %s and reads it as the DOM does', (_, edit) => {
		const plain = sheet(`xmlns="${NS.x}" xmlns:x14ac="urn:x14ac"`, CELLS);
		const xml = edit(plain);
		expect(xml).not.toBe(plain);
		const { split, dom, lite } = readBothWays(xml);
		expect(split).toBeDefined();
		expect(lite).toEqual(dom);
		// Everything outside the cell data stays in the part the DOM parses.
		const rest = parseXml(split?.xml ?? '').documentElement;
		expect(elements(rest).map((node) => node.localName)).toEqual(['sheetData', 'pageMargins']);
		expect(elements(elements(rest)[0]!)).toHaveLength(0);
	});

	it.each([
		[
			'a processing instruction',
			(xml: string) => xml.replace('<row r="3">', '<?keep?><row r="3">'),
		],
		[
			'a comment XML forbids',
			(xml: string) => xml.replace('<row r="3">', '<!-- a -- b --><row r="3">'),
		],
		['an unterminated comment', (xml: string) => xml.replace('<row r="3">', '<!-- c <row r="3">')],
		['unterminated CDATA', (xml: string) => xml.replace('<t>Bold</t>', '<t><![CDATA[Bold</t>')],
		[
			'CDATA quoting a DTD',
			(xml: string) => xml.replace('<t>Bold</t>', '<t><![CDATA[<!doctype html>]]></t>'),
		],
		[
			'a comment quoting an entity declaration',
			(xml: string) => xml.replace('<row r="3">', '<!-- <!ENTITY x "y"> --><row r="3">'),
		],
		[
			'an undeclared prefix',
			(xml: string) =>
				xml
					.replace('<c r="B4"', '<y:c r="B4"')
					.replace('</c></row><row r="5">', '</y:c></row><row r="5">'),
		],
		[
			'a self-closing sheetData',
			(xml: string) => xml.replace(/<sheetData>[^]*<\/sheetData>/, '<sheetData/>'),
		],
		['no sheetData', (xml: string) => xml.replace(/<sheetData>[^]*<\/sheetData>/, '')],
	])('leaves a sheet with %s to the DOM parser', (_, edit) => {
		expect(
			splitSheetData(edit(sheet(`xmlns="${NS.x}" xmlns:x14ac="urn:x14ac"`, CELLS))),
		).toBeUndefined();
	});

	it('still refuses a part whose cell data quotes a DTD or entity declaration', () => {
		const base = sheet(`xmlns="${NS.x}" xmlns:x14ac="urn:x14ac"`, CELLS);
		for (const xml of [
			base.replace('<t>Bold</t>', '<t><![CDATA[<!DOCTYPE html>]]></t>'),
			base.replace('<row r="3">', '<!-- <!entity x "y"> --><row r="3">'),
		]) {
			expect(splitSheetData(xml)).toBeUndefined();
			expect(() => parseXml(xml)).toThrow(/DTD or entity declarations/);
		}
	});
});
