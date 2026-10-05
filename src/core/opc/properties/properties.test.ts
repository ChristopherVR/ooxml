import { describe, expect, it } from 'vitest';
import {
	allocateCustomPropertyIds,
	parseAppProperties,
	parseCoreProperties,
	parseCustomProperties,
	writeAppProperties,
	writeCoreProperties,
	writeCustomProperties,
	type CoreProperties,
	type CustomProperty,
} from './index.js';

const EXCEL_CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Budget</dc:title><dc:creator>Ann</dc:creator><cp:lastModifiedBy>Bob</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-02T03:04:05Z</dcterms:modified><cp:category>Finance</cp:category><cp:contentStatus>Draft</cp:contentStatus><dcterms:extra>keep me</dcterms:extra></cp:coreProperties>`;

const EXCEL_APP = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></TitlesOfParts><Company>Contoso</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion><DigSig><vt:blob>AAAA</vt:blob></DigSig></Properties>`;

describe('core properties', () => {
	it('reads every Dublin Core and cp field', () => {
		expect(parseCoreProperties(EXCEL_CORE)).toEqual({
			title: 'Budget',
			creator: 'Ann',
			lastModifiedBy: 'Bob',
			created: '2026-01-02T03:04:05Z',
			modified: '2026-02-02T03:04:05Z',
			category: 'Finance',
			contentStatus: 'Draft',
		});
		expect(parseCoreProperties(undefined)).toEqual({});
	});

	it('writes all fields from scratch and reads them back', () => {
		const all: Required<CoreProperties> = {
			title: 'T',
			subject: 'S',
			creator: 'C',
			keywords: 'k1; k2',
			description: 'D & <d>',
			lastModifiedBy: 'L',
			created: '2026-01-01T00:00:00Z',
			modified: '2026-01-02T00:00:00Z',
			category: 'Cat',
			contentStatus: 'Final',
			revision: '7',
			language: 'en-ZA',
			identifier: 'urn:x',
			lastPrinted: '2026-01-03T00:00:00Z',
			version: '1.2',
		};
		const xml = writeCoreProperties(all);
		expect(xml.startsWith('<?xml')).toBe(true);
		expect(xml).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">');
		expect(parseCoreProperties(xml)).toEqual(all);
	});

	it('patches the source part, keeping unknown elements and removing cleared fields', () => {
		const props = parseCoreProperties(EXCEL_CORE);
		delete props.category;
		const xml = writeCoreProperties({ ...props, title: 'New', revision: '3' }, EXCEL_CORE);
		expect(xml).toContain('<dcterms:extra>keep me</dcterms:extra>');
		expect(xml).not.toContain('Finance');
		expect(parseCoreProperties(xml)).toMatchObject({ title: 'New', revision: '3', creator: 'Ann' });
	});
});

describe('app properties', () => {
	it('reads known fields including heading pairs and titles', () => {
		expect(parseAppProperties(EXCEL_APP)).toEqual({
			application: 'Microsoft Excel',
			docSecurity: 0,
			scaleCrop: false,
			headingPairs: [
				{ name: 'Worksheets', count: 2 },
				{ name: 'Named Ranges', count: 1 },
			],
			titlesOfParts: ['Sheet1', 'Data', 'Data!Print_Area'],
			company: 'Contoso',
			linksUpToDate: false,
			sharedDoc: false,
			hyperlinksChanged: false,
			appVersion: '16.0300',
		});
	});

	it('leaves the source byte for byte when nothing changed', () => {
		const xml = writeAppProperties(parseAppProperties(EXCEL_APP), EXCEL_APP);
		expect(xml.replace(/\r?\n/g, '')).toBe(EXCEL_APP.replace(/\r?\n/g, ''));
	});

	it('patches changed fields and keeps unknown elements', () => {
		const props = parseAppProperties(EXCEL_APP);
		const xml = writeAppProperties(
			{
				...props,
				manager: 'Mia',
				hyperlinkBase: 'https://example.com/',
				titlesOfParts: ['Renamed', 'Data', 'Data!Print_Area'],
			},
			EXCEL_APP,
		);
		expect(xml).toContain('<DigSig><vt:blob>AAAA</vt:blob></DigSig>');
		const back = parseAppProperties(xml);
		expect(back.manager).toBe('Mia');
		expect(back.hyperlinkBase).toBe('https://example.com/');
		expect(back.titlesOfParts).toEqual(['Renamed', 'Data', 'Data!Print_Area']);
		expect(back.headingPairs).toEqual(props.headingPairs);
		// Manager goes before Company, as Excel orders it.
		expect(xml.indexOf('<Manager>')).toBeLessThan(xml.indexOf('<Company>'));
	});

	it('writes a fresh part', () => {
		const xml = writeAppProperties({
			application: 'Microsoft Excel',
			docSecurity: 2,
			headingPairs: [{ name: 'Worksheets', count: 1 }],
			titlesOfParts: ['A'],
		});
		expect(xml).toContain('<vt:vector size="2" baseType="variant">');
		expect(parseAppProperties(xml)).toEqual({
			application: 'Microsoft Excel',
			docSecurity: 2,
			headingPairs: [{ name: 'Worksheets', count: 1 }],
			titlesOfParts: ['A'],
		});
	});
});

describe('custom properties', () => {
	const props: CustomProperty[] = [
		{ name: 'Project', type: 'lpwstr', value: 'Apollo & <Co>' },
		{ name: 'Count', type: 'i4', value: 42 },
		{ name: 'Ratio', type: 'r8', value: 0.25 },
		{ name: 'Approved', type: 'bool', value: true },
		{ name: 'Due', type: 'filetime', value: '2026-10-03T00:00:00Z' },
	];

	it('round-trips every editable variant type with allocated pids', () => {
		const xml = writeCustomProperties(props);
		expect(xml).toBeDefined();
		const back = parseCustomProperties(xml);
		expect(back).toEqual(props.map((p, index) => ({ ...p, pid: index + 2 })));
		expect(xml).toContain('fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"');
	});

	it('keeps unknown variants raw and link targets', () => {
		const xml = `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Big"><vt:i8>9007199254740993</vt:i8></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Linked" linkTarget="MyName"><vt:lpwstr>x</vt:lpwstr></property></Properties>`;
		const back = parseCustomProperties(xml);
		expect(back[0]).toMatchObject({ name: 'Big', type: 'raw', pid: 5 });
		expect(back[1]).toEqual({
			name: 'Linked',
			type: 'lpwstr',
			value: 'x',
			pid: 3,
			linkTarget: 'MyName',
		});
		const again = writeCustomProperties(back) ?? '';
		expect(again).toContain('9007199254740993</vt:i8>');
		expect(again).toContain('linkTarget="MyName"');
		expect(parseCustomProperties(again)).toEqual(back);
	});

	it('allocates ids after the largest kept one and drops duplicate names', () => {
		expect(
			allocateCustomPropertyIds([
				{ name: 'a', type: 'bool', value: true, pid: 7 },
				{ name: 'b', type: 'bool', value: true },
				{ name: 'c', type: 'bool', value: true, pid: 7 },
				{ name: 'd', type: 'bool', value: true, pid: 1 },
			]),
		).toEqual([7, 8, 9, 10]);
		const xml = writeCustomProperties([
			{ name: 'A', type: 'i4', value: 1 },
			{ name: 'a', type: 'i4', value: 2 },
		]);
		expect(parseCustomProperties(xml)).toHaveLength(1);
		expect(writeCustomProperties([])).toBeUndefined();
	});
});
