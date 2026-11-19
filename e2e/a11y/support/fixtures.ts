/**
 * The documents the gates open. Word and Excel are built here so every accessibility property they
 * carry is visible in one place: a title, two heading levels, alt text, a table with a repeated
 * header row, caption and description, run language, a right-to-left paragraph and a link. The
 * PowerPoint and Visio files are committed fixtures of their own browser suites.
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

export interface FixtureFile {
	name: string;
	mimeType: string;
	buffer: Buffer;
}

const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
	'base64',
);
const OCTET = 'application/octet-stream';

async function zip(name: string, files: Record<string, string | Buffer>): Promise<FixtureFile> {
	const archive = new JSZip();
	for (const [path, data] of Object.entries(files)) archive.file(path, data);
	return { name, mimeType: OCTET, buffer: await archive.generateAsync({ type: 'nodebuffer' }) };
}

/** What the Word fixture holds, for the tree and export checks. */
export const WORD = {
	title: 'Annual report',
	heading1: 'First section',
	heading2: 'Sub section',
	alt: 'Bar chart of quarterly sales',
	header: ['Quarter', 'Sales'],
	caption: 'Quarter table',
	language: 'en-GB',
	french: 'Bonjour le monde',
	link: 'Example link',
} as const;

export function wordDocument(): Promise<FixtureFile> {
	const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
	const p = (text: string, style?: string) =>
		`<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
	const image = `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1270000" cy="1270000"/><wp:docPr id="7" name="Picture 7" title="Sales title" descr="${WORD.alt}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="p.png" descr="${WORD.alt}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
	const cell = (text: string) =>
		`<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
	const table = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblCaption w:val="${WORD.caption}"/><w:tblDescription w:val="Sales per quarter"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>${WORD.header.map(cell).join('')}</w:tr><w:tr>${cell('Q1')}${cell('100')}</w:tr><w:tr>${cell('Q2')}${cell('200')}</w:tr></w:tbl>`;
	const body = [
		p(WORD.title, 'Title'),
		p(WORD.heading1, 'Heading1'),
		p('Body paragraph alpha with a sentence for reading.'),
		p(WORD.heading2, 'Heading2'),
		image,
		table,
		p('Second body paragraph bravo.'),
		'<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>مرحبا بالعالم</w:t></w:r></w:p>',
		`<w:p><w:r><w:rPr><w:lang w:val="fr-FR"/></w:rPr><w:t>${WORD.french}</w:t></w:r></w:p>`,
		`<w:p><w:hyperlink r:id="rId11"><w:r><w:t>${WORD.link}</w:t></w:r></w:hyperlink></w:p>`,
	].join('');
	const style = (id: string, name: string, extra: string) =>
		`<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/>${extra}</w:style>`;
	const styles = `${X}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="${WORD.language}" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${style('Title', 'Title', '<w:rPr><w:sz w:val="48"/></w:rPr>')}${style('Heading1', 'heading 1', '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr>')}${style('Heading2', 'heading 2', '<w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr>')}</w:styles>`;
	return zip('Accessible report.docx', {
		'[Content_Types].xml': `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`,
		'_rels/.rels': `${X}<Relationships ${PKG}><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
		'docProps/core.xml': `${X}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Annual report title</dc:title><dc:language>${WORD.language}</dc:language></cp:coreProperties>`,
		'word/document.xml': `${X}<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
		'word/styles.xml': styles,
		'word/_rels/document.xml.rels': `${X}<Relationships ${PKG}><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId10" Type="${REL}/image" Target="media/p.png"/><Relationship Id="rId11" Type="${REL}/hyperlink" Target="https://example.invalid/" TargetMode="External"/></Relationships>`,
		'word/media/p.png': PNG,
	});
}

/** What the Excel fixture holds. */
export const EXCEL = {
	sheets: ['Sales', 'Notes'],
	header: ['Quarter', 'Sales'],
	tableAlt: 'Sales table alt',
	tableSummary: 'Quarterly sales',
	pictureAlt: 'Sales picture alt',
} as const;

export function excelWorkbook(): Promise<FixtureFile> {
	const S = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"`;
	const row = (r: number, cells: [string, string | number][]) =>
		`<row r="${r}">${cells.map(([a, v]) => (typeof v === 'number' ? `<c r="${a}"><v>${v}</v></c>` : `<c r="${a}" t="inlineStr"><is><t>${v}</t></is></c>`)).join('')}</row>`;
	const type = 'application/vnd.openxmlformats-officedocument';
	return zip('Accessible sales.xlsx', {
		'[Content_Types].xml': `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="${type}.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${type}.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="${type}.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="${type}.spreadsheetml.table+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="${type}.drawing+xml"/></Types>`,
		'_rels/.rels': `${X}<Relationships ${PKG}><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		'xl/workbook.xml': `${X}<workbook ${S}><sheets><sheet name="${EXCEL.sheets[0]}" sheetId="1" r:id="rId1"/><sheet name="${EXCEL.sheets[1]}" sheetId="2" r:id="rId2"/></sheets></workbook>`,
		'xl/_rels/workbook.xml.rels': `${X}<Relationships ${PKG}><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
		'xl/worksheets/sheet1.xml': `${X}<worksheet ${S}><sheetData>${row(1, [
			['A1', 'Quarter'],
			['B1', 'Sales'],
		])}${row(2, [
			['A2', 'Q1'],
			['B2', 100],
		])}${row(3, [
			['A3', 'Q2'],
			['B3', 200],
		])}${row(4, [
			['A4', 'Q3'],
			['B4', 300],
		])}</sheetData><drawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `${X}<Relationships ${PKG}><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
		'xl/worksheets/sheet2.xml': `${X}<worksheet ${S}><sheetData>${row(1, [['A1', 'Note']])}</sheetData></worksheet>`,
		'xl/tables/table1.xml': `${X}<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:B4" headerRowCount="1" altText="${EXCEL.tableAlt}" altTextSummary="${EXCEL.tableSummary}"><autoFilter ref="A1:B4"/><tableColumns count="2"><tableColumn id="1" name="Quarter"/><tableColumn id="2" name="Sales"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`,
		'xl/drawings/drawing1.xml': `${X}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}"><xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="${EXCEL.pictureAlt}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`,
		'xl/drawings/_rels/drawing1.xml.rels': `${X}<Relationships ${PKG}><Relationship Id="rId1" Type="${REL}/image" Target="../media/p.png"/></Relationships>`,
		'xl/media/p.png': PNG,
	});
}

const committed = (path: string, name: string): FixtureFile => ({
	name,
	mimeType: OCTET,
	buffer: readFileSync(fileURLToPath(new URL(`../../../${path}`, import.meta.url))),
});

/** What the PowerPoint fixture holds: one slide, a described picture and a decorative one. */
export const POWERPOINT = { alt: 'A wireframe globe icon' } as const;

export const powerPointDeck = (): FixtureFile =>
	committed('e2e/pptx/fixtures/accessibility-images.pptx', 'Accessible deck.pptx');

export const visioDrawing = (): FixtureFile =>
	committed('src/core/visio/__fixtures__/dynamic-connector-native.vsdx', 'Process.vsdx');

/** A one-slide deck whose two shapes hold text, for edits that change a run. */
export const powerPointTextDeck = (): FixtureFile =>
	committed('e2e/pptx/fixtures/format-painter.pptx', 'Deck.pptx');
