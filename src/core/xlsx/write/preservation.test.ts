// The preservation guarantee of viewers/xlsx/docs/preservation.md: parts, relationships and XML the
// model does not represent survive an edit and a save, or are reported. The package below holds
// one of each kind of unmodelled content; an edit changes one cell and the save must keep the rest.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/index';
import { createEditSession } from '../edit/session';
import { saveXlsx } from './index';
import { HEAD, R, X, miniPackage } from './mini-package-fixtures';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const bytes = (...values: number[]) => new Uint8Array(values);

/** Source parts that must come back byte for byte. */
const VERBATIM: Record<string, string | Uint8Array> = {
	'xl/vbaProject.bin': bytes(0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4, 5),
	'xl/printerSettings/printerSettings1.bin': bytes(9, 8, 7, 6),
	'customXml/item1.xml': `${HEAD}<root xmlns="urn:custom"><value>kept</value></root>`,
	'customXml/itemProps1.xml': `${HEAD}<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{11111111-2222-3333-4444-555555555555}"/>`,
	'customXml/_rels/item1.xml.rels': `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
	'xl/externalLinks/externalLink1.xml': `${HEAD}<externalLink xmlns="${X}"><externalBook xmlns:r="${R}" r:id="rId1"><sheetNames><sheetName val="Other"/></sheetNames></externalBook></externalLink>`,
	'xl/externalLinks/_rels/externalLink1.xml.rels': `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/externalLinkPath" Target="file:///C:/data/other.xlsx" TargetMode="External"/></Relationships>`,
	'xl/connections.xml': `${HEAD}<connections xmlns="${X}"><connection id="1" name="Conn" type="1"/></connections>`,
	'xl/ctrlProps/ctrlProp1.xml': `${HEAD}<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox"/>`,
	'xl/custom/blob.dat': bytes(1, 1, 2, 3, 5, 8),
};

const WORKBOOK = `${HEAD}<workbook xmlns="${X}" xmlns:r="${R}" xmlns:mc="${MC}" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac">
<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27328"/>
<fileSharing readOnlyRecommended="1" userName="author"/>
<workbookPr defaultThemeVersion="166925" codeName="ThisWorkbook"/>
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="9000" tabRatio="700"/><workbookView xWindow="100" yWindow="100" windowWidth="10000" windowHeight="5000"/></bookViews>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<functionGroups builtInGroupCount="2"/>
<externalReferences><externalReference r:id="rIdExt"/></externalReferences>
<definedNames><definedName name="Macro1" function="1" vbProcedure="1" description="A macro" customMenu="Menu" help="Help" statusBar="Status" shortcutKey="m">Sheet1!$A$1</definedName><definedName name="Plain">Sheet1!$B$1</definedName></definedNames>
<calcPr calcId="191029" iterate="1" iterateCount="50"/>
<oleSize ref="A1:B2"/>
<customWorkbookViews><customWorkbookView name="Mine" guid="{AAAAAAAA-0000-0000-0000-000000000001}" windowWidth="800" windowHeight="600" activeSheetId="1"/></customWorkbookViews>
<smartTagPr embed="1"/>
<webPublishing css="1"/>
<fileRecoveryPr autoRecover="0"/>
<extLst><ext uri="{00000000-0000-0000-0000-00000000ABCD}" xmlns:u="urn:unknown"><u:setting name="keep">me</u:setting></ext></extLst>
</workbook>`;

const SHEET = `${HEAD}<worksheet xmlns="${X}" xmlns:r="${R}" xmlns:mc="${MC}" xmlns:u="urn:unknown">
<sheetPr codeName="Sheet1"><tabColor rgb="FF00FF00"/><pageSetUpPr fitToPage="1"/></sheetPr>
<dimension ref="A1:B2"/>
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<sheetData><row r="1" spans="1:2" thickBot="1" thickTop="1"><c r="A1"><v>1</v></c><c r="B1" cm="1"><v>2</v></c></row><row r="2"><c r="A2" ph="1" t="inlineStr"><is><t>x</t></is></c></row></sheetData>
<sheetCalcPr fullCalcOnLoad="1"/>
<phoneticPr fontId="1" type="noConversion"/>
<printOptions horizontalCentered="1"/>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
<pageSetup paperSize="9" orientation="landscape" r:id="rId1"/>
<rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"/></rowBreaks>
<colBreaks count="1" manualBreakCount="1"><brk id="5" max="1048575" man="1"/></colBreaks>
<customProperties><customPr name="tag" r:id="rId2"/></customProperties>
<cellWatches><cellWatch r="A1"/></cellWatches>
<ignoredErrors><ignoredError sqref="A1" numberStoredAsText="1"/></ignoredErrors>
<smartTags><cellSmartTags r="A1"><cellSmartTag type="0"/></cellSmartTags></smartTags>
<controls><mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="1025" r:id="rId3" name="Check Box 1"/></mc:Choice></mc:AlternateContent></controls>
<extLst><ext uri="{00000000-0000-0000-0000-00000000DCBA}"><u:payload>kept</u:payload></ext></extLst>
</worksheet>`;

const SHEET_RELS = `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/printerSettings" Target="../printerSettings/printerSettings1.bin"/><Relationship Id="rId2" Type="${REL}/customProperty" Target="../custom/blob.dat"/><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>`;

const WORKBOOK_RELS =
	`<Relationship Id="rIdV" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>` +
	`<Relationship Id="rIdC" Type="${REL}/customXml" Target="../customXml/item1.xml"/>` +
	`<Relationship Id="rIdExt" Type="${REL}/externalLink" Target="externalLinks/externalLink1.xml"/>` +
	`<Relationship Id="rIdN" Type="${REL}/connections" Target="connections.xml"/>`;

const CONTENT_TYPES =
	'<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' +
	'<Default Extension="dat" ContentType="application/octet-stream"/>' +
	'<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>' +
	'<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
	'<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>' +
	'<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>' +
	'<Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>';

async function buildSource(): Promise<Uint8Array> {
	const base = await miniPackage({
		sheets: [{ name: 'Sheet1', xml: SHEET, rels: SHEET_RELS }],
		workbookRels: WORKBOOK_RELS,
		contentTypes: CONTENT_TYPES,
		parts: { 'xl/workbook.xml': WORKBOOK, ...(VERBATIM as Record<string, string>) },
	});
	const zip = await JSZip.loadAsync(base);
	// A macro-enabled workbook: VBA is kept only for .xlsm.
	const types = await zip.file('[Content_Types].xml')!.async('string');
	zip.file(
		'[Content_Types].xml',
		types.replace(
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
			'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
		),
	);
	for (const [name, data] of Object.entries(VERBATIM)) zip.file(name, data);
	return zip.generateAsync({ type: 'uint8array' });
}

async function editAndSave(source: Uint8Array) {
	const wb = await loadXlsx(source);
	const session = createEditSession(wb);
	session.setCellValue(0, 3, 0, 'edited');
	const saved = await saveXlsx(wb);
	const zip = await JSZip.loadAsync(saved);
	return { wb, saved, zip, text: async (name: string) => zip.file(name)?.async('string') };
}

describe('preservation of unmodelled content through an edit and a save', () => {
	it('keeps every unmodelled part byte for byte', async () => {
		const source = await JSZip.loadAsync(await buildSource());
		const { zip } = await editAndSave(await buildSource());
		for (const name of Object.keys(VERBATIM)) {
			const before = await source.file(name)!.async('uint8array');
			const after = await zip.file(name)?.async('uint8array');
			expect(after, name).toBeDefined();
			expect(Array.from(after!), name).toEqual(Array.from(before));
		}
		// Their own relationships (an external link target, custom XML properties) are kept too.
		for (const name of [
			'customXml/_rels/item1.xml.rels',
			'xl/externalLinks/_rels/externalLink1.xml.rels',
		])
			expect(await zip.file(name)?.async('string'), name).toBe(
				await source.file(name)!.async('string'),
			);
	});

	it('keeps the workbook relationships and content types of the carried parts', async () => {
		const { text } = await editAndSave(await buildSource());
		const rels = (await text('xl/_rels/workbook.xml.rels'))!;
		for (const target of [
			'vbaProject.bin',
			'../customXml/item1.xml',
			'externalLinks/externalLink1.xml',
			'connections.xml',
		])
			expect(rels, target).toContain(`Target="${target}"`);
		const types = (await text('[Content_Types].xml'))!;
		expect(types).toContain('macroEnabled.main+xml');
		for (const part of [
			'/customXml/itemProps1.xml',
			'/xl/externalLinks/externalLink1.xml',
			'/xl/connections.xml',
			'/xl/ctrlProps/ctrlProp1.xml',
		])
			expect(types, part).toContain(`PartName="${part}"`);
		const sheetRels = (await text('xl/worksheets/_rels/sheet1.xml.rels'))!;
		for (const target of [
			'../printerSettings/printerSettings1.bin',
			'../custom/blob.dat',
			'../ctrlProps/ctrlProp1.xml',
		])
			expect(sheetRels, target).toContain(`Target="${target}"`);
	});

	it('keeps the workbook elements the model does not represent', async () => {
		const { text } = await editAndSave(await buildSource());
		const xml = (await text('xl/workbook.xml'))!;
		for (const fragment of [
			'<fileVersion appName="xl"',
			'readOnlyRecommended="1"',
			'codeName="ThisWorkbook"',
			'<functionGroups builtInGroupCount="2"',
			'r:id="rIdExt"',
			'<oleSize ref="A1:B2"',
			'<customWorkbookView name="Mine"',
			'<smartTagPr embed="1"',
			'<webPublishing css="1"',
			'<fileRecoveryPr autoRecover="0"',
			'urn:unknown',
			'name="keep">me<',
		])
			expect(xml, fragment).toContain(fragment);
	});

	it('keeps every window of the workbook and the attributes of defined names', async () => {
		const { text } = await editAndSave(await buildSource());
		const xml = (await text('xl/workbook.xml'))!;
		expect(xml.match(/<workbookView\b/g)).toHaveLength(2);
		expect(xml).toContain('tabRatio="700"');
		for (const attribute of [
			'function="1"',
			'vbProcedure="1"',
			'description="A macro"',
			'customMenu="Menu"',
			'help="Help"',
			'statusBar="Status"',
			'shortcutKey="m"',
		])
			expect(xml, attribute).toContain(attribute);
	});

	it('keeps the worksheet elements the model does not represent, in schema order', async () => {
		const { text } = await editAndSave(await buildSource());
		const xml = (await text('xl/worksheets/sheet1.xml'))!;
		const order = [
			'<sheetCalcPr',
			'<phoneticPr',
			'<printOptions',
			'<pageSetup',
			'<rowBreaks',
			'<colBreaks',
			'<customProperties',
			'<cellWatches',
			'<ignoredErrors',
			'<smartTags',
			'<controls',
			'<extLst',
		];
		const positions = order.map((tag) => xml.indexOf(tag));
		expect(positions.map((p, i) => [order[i], p >= 0])).toEqual(order.map((tag) => [tag, true]));
		expect([...positions].sort((a, b) => a - b)).toEqual(positions);
		expect(xml).toContain('codeName="Sheet1"');
		expect(xml).toContain('r:id="rId1"');
		expect(xml).toContain('>kept</u:payload>');
	});

	it('writes the edit and nothing else changed in the cells', async () => {
		const { text, wb } = await editAndSave(await buildSource());
		expect(await text('xl/sharedStrings.xml')).toContain('edited');
		expect(wb.sheets[0]?.name).toBe('Sheet1');
	});

	it('saves a second time to the same bytes for the carried parts', async () => {
		const first = await editAndSave(await buildSource());
		const second = await JSZip.loadAsync(await saveXlsx(await loadXlsx(first.saved)));
		for (const name of Object.keys(VERBATIM))
			expect(Array.from((await second.file(name)?.async('uint8array')) ?? []), name).toEqual(
				Array.from((await first.zip.file(name)!.async('uint8array')) ?? []),
			);
	});

	it('keeps row hints and the attributes of macro and function names across a reload', async () => {
		const { saved } = await editAndSave(await buildSource());
		const again = await loadXlsx(saved);
		expect(again.sheets[0]?.rowInfo.get(0)).toMatchObject({ thickTop: true, thickBot: true });
		expect(again.definedNames.find((n) => n.name === 'Macro1')?.extra).toMatchObject({
			function: '1',
			vbProcedure: '1',
			shortcutKey: 'm',
		});
		expect(again.definedNames.find((n) => n.name === 'Plain')?.extra).toBeUndefined();
	});

	it('reports what it keeps without running or refreshing', async () => {
		const wb = await loadXlsx(await buildSource());
		expect(wb.warnings).toContain(
			'The workbook contains VBA macros; they are kept on save but never run.',
		);
		expect(wb.warnings).toContain('External workbook links are kept but not updated.');
	});

	it('reports a worksheet element outside the schema, which a save cannot place', async () => {
		const bytes = await miniPackage({
			sheets: [
				{
					name: 'S',
					xml: `${HEAD}<worksheet xmlns="${X}" xmlns:f="urn:future"><sheetData/><f:futureThing a="1"/></worksheet>`,
				},
			],
		});
		const wb = await loadXlsx(bytes);
		expect(wb.warnings.join(' ')).toContain('"futureThing" is not part of the worksheet schema');
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await saved.file('xl/worksheets/sheet1.xml')!.async('string')).not.toContain(
			'futureThing',
		);
	});
});
