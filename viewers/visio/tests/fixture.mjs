import JSZip from 'jszip';

/** Original synthetic OPC fixture; no copied Microsoft diagram or user document. */
export async function createVsdxFixture(text = 'Imported from VSDX') {
	const zip = new JSZip();
	const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
	const visio = 'http://schemas.microsoft.com/visio/2010/relationships/';
	const ns = 'http://schemas.microsoft.com/office/visio/2012/main';
	const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	const escape = (value) =>
		value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${visio}document" Target="visio/document.xml"/></Relationships>`,
	);
	zip.file('visio/document.xml', `<VisioDocument xmlns="${ns}"/>`);
	zip.file(
		'visio/_rels/document.xml.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${visio}pages" Target="pages/pages.xml"/></Relationships>`,
	);
	zip.file(
		'visio/pages/pages.xml',
		`<Pages xmlns="${ns}" xmlns:r="${relNs}"><Page ID="1" Name="Imported page"><PageSheet><Cell N="PageWidth" V="8.5"/><Cell N="PageHeight" V="11"/></PageSheet><Rel r:id="rId1"/></Page></Pages>`,
	);
	zip.file(
		'visio/pages/_rels/pages.xml.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${visio}page" Target="page1.xml"/></Relationships>`,
	);
	zip.file(
		'visio/pages/page1.xml',
		`<PageContents xmlns="${ns}"><Shapes><Shape ID="1" NameU="Import test"><Cell N="Width" V="3"/><Cell N="Height" V="1"/><Cell N="PinX" V="4"/><Cell N="PinY" V="7"/><Cell N="FillForegnd" V="#DAEFE6"/><Section N="Geometry" IX="0"><Row T="MoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="2"><Cell N="X" V="3"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="3"><Cell N="X" V="3"/><Cell N="Y" V="1"/></Row><Row T="LineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="1"/></Row><Row T="LineTo" IX="5"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row></Section><Text>${escape(text)}</Text></Shape></Shapes></PageContents>`,
	);
	return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
