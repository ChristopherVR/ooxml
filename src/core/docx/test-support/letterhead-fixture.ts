import JSZip from 'jszip';

// A letterhead shaped like the ones Word saves: a first-page header (`w:titlePg`,
// `w:headerReference w:type="first"`) holding a page-anchored text box written as a plain `rect`
// with a `wps:txbx` (no `txBox="1"`), and a `straightConnector1` line beside the body.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NAMESPACES = [
	`xmlns:w="${W}"`,
	`xmlns:r="${R}"`,
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
	'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
].join(' ');

const anchor = (id: number, name: string, x: number, y: number, cx: number, cy: number) =>
	`<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="25166${id}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">` +
	`<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
	`<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
	`<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
	`<wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/>`;

/** The partners' text box: bold names, size-9 details and a paragraph with a bottom border. */
export const LETTERHEAD_TEXT_BOX =
	'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>' +
	anchor(11, 'Text Box 6', 323215, 2189480, 2032000, 8183244) +
	'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>' +
	'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2032000" cy="8183244"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
	'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="6350"><a:noFill/></a:ln></wps:spPr>' +
	'<wps:txbx><w:txbxContent>' +
	'<w:p><w:r><w:rPr><w:b/><w:color w:val="1F3864"/></w:rPr><w:t>First Partner</w:t></w:r></w:p>' +
	'<w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>Specialist in contract law</w:t></w:r></w:p>' +
	'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="1F3864"/></w:pBdr></w:pPr></w:p>' +
	'<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="18"/></w:rPr><w:t>12 Example Street</w:t></w:r></w:p>' +
	'</w:txbxContent></wps:txbx>' +
	'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>' +
	'</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

/** The vertical rule next to the body: a `straightConnector1` with a dashed blue outline. */
export const LETTERHEAD_CONNECTOR =
	'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>' +
	anchor(12, 'Straight Connector 7', 2450000, 2189480, 0, 8183244) +
	'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/>' +
	'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="8183244"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>' +
	'<a:ln w="19050"><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill><a:prstDash val="dash"/></a:ln></wps:spPr>' +
	'<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

/** The first-page header part's XML. */
export const LETTERHEAD_HEADER_XML = `<w:hdr ${NAMESPACES}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>${LETTERHEAD_TEXT_BOX}${LETTERHEAD_CONNECTOR}<w:r><w:t>Letterhead</w:t></w:r></w:p></w:hdr>`;

/** A one-page document whose first-page header holds the letterhead text box and rule. */
export async function letterheadFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<?xml version="1.0" encoding="UTF-8"?><w:document ${NAMESPACES}><w:body><w:p><w:r><w:t>Dear client,</w:t></w:r></w:p><w:p><w:r><w:t>The letter body starts here.</w:t></w:r></w:p>` +
			`<w:sectPr><w:headerReference w:type="first" r:id="rId1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="3400" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/></Relationships>`,
	);
	zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8"?>${LETTERHEAD_HEADER_XML}`);
	return zip.generateAsync({ type: 'uint8array' });
}
