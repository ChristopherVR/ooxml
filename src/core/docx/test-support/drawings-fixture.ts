import JSZip from 'jszip';
import { LETTERHEAD_CONNECTOR, LETTERHEAD_TEXT_BOX } from './letterhead-fixture';

// The letterhead with every kind of drawing a header or body holds, shaped like Word's files: an
// inline logo picture in the first-page header next to the page-anchored text box and rule; in the
// body, a floating and an inline text box written with `txBox="1"`, and a shape whose text colour
// comes from its `wps:style` font reference.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NAMESPACES = [
	`xmlns:w="${W}"`,
	`xmlns:r="${R}"`,
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(' ');
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

/** An 8 x 4 navy PNG. */
export const LOGO_PNG = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAEUlEQVR4nGOQt0jBihioJwEAiMAXYQzn07UAAAAASUVORK5CYII=',
		'base64',
	),
);

/** The header's inline logo (`wp:inline`), 96 x 48 px. */
const LOGO =
	'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
	'<wp:docPr id="10" name="Picture 10" descr="Logo"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
	'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
	'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
	'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

const wsp = (cNvSpPr: string, cx: number, cy: number, spPr: string, content: string, style = '') =>
	`<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp>${cNvSpPr}<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
	`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPr}</wps:spPr>${style}<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
	'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>';

/** A floating text box Word's Insert > Text Box writes: `txBox="1"`, square wrap, red outline. */
export const FLOATING_TEXT_BOX =
	'<w:r><w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
	'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>2743200</wp:posOffset></wp:positionH>' +
	'<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
	'<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="21" name="Text Box 21"/><wp:cNvGraphicFramePr/>' +
	wsp(
		'<wps:cNvSpPr txBox="1"/>',
		1828800,
		914400,
		'<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>',
		'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Floating note</w:t></w:r></w:p><w:p><w:r><w:t>Anchored text box</w:t></w:r></w:p>',
	) +
	'</wp:anchor></w:drawing></w:r>';

/** An inline text box in the simple form the editor edits (plain runs, `txBox="1"`). */
export const INLINE_TEXT_BOX =
	'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1371600" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
	'<wp:docPr id="22" name="Text Box 22"/><wp:cNvGraphicFramePr/>' +
	wsp(
		'<wps:cNvSpPr txBox="1"/>',
		1371600,
		457200,
		'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
		'<w:p><w:r><w:t>Inline box</w:t></w:r></w:p>',
	) +
	'</wp:inline></w:drawing></w:r>';

/**
 * A light yellow shape whose text colour comes from its style's font reference (`lt1` at half
 * luminance, a grey), which Word uses rather than black on the light fill.
 */
export const STYLED_SHAPE =
	'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1371600" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
	'<wp:docPr id="23" name="Rectangle 23"/><wp:cNvGraphicFramePr/>' +
	wsp(
		'<wps:cNvSpPr/>',
		1371600,
		457200,
		'<a:solidFill><a:srgbClr val="FFFF66"/></a:solidFill>',
		'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Styled text</w:t></w:r></w:p>',
		'<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
			'<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"><a:lumMod val="50000"/></a:schemeClr></a:fontRef></wps:style>',
	) +
	'</wp:inline></w:drawing></w:r>';

/** The first-page header part: the logo inline, then the letterhead's floating box and rule. */
export const DRAWINGS_HEADER_XML = `<w:hdr ${NAMESPACES}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>${LOGO}${LETTERHEAD_TEXT_BOX}${LETTERHEAD_CONNECTOR}<w:r><w:t>Letterhead</w:t></w:r></w:p></w:hdr>`;

const PH_FILL = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';

/** Word's default Office colour scheme, which the style's `lt1` resolves through. */
const THEME_XML =
	'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office">' +
	'<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
	'<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
	'<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
	'<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
	'<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
	'<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
	// Word refuses a theme without its format scheme: three fills, outlines, effects and backgrounds.
	`<a:fmtScheme name="Office"><a:fillStyleLst>${PH_FILL.repeat(3)}</a:fillStyleLst><a:lnStyleLst>` +
	[6350, 12700, 19050].map((w) => `<a:ln w="${w}">${PH_FILL}</a:ln>`).join('') +
	'</a:lnStyleLst><a:effectStyleLst>' +
	'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3) +
	`</a:effectStyleLst><a:bgFillStyleLst>${PH_FILL.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

/** The letterhead with a header logo and the body's text boxes and styled shape. */
export async function drawingsFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<?xml version="1.0" encoding="UTF-8"?><w:document ${NAMESPACES}><w:body>` +
			`<w:p>${FLOATING_TEXT_BOX}<w:r><w:t>Dear client, the letter body starts here.</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>Inline: </w:t></w:r>${INLINE_TEXT_BOX}<w:r><w:t> and </w:t></w:r>${STYLED_SHAPE}</w:p>` +
			`<w:sectPr><w:headerReference w:type="first" r:id="rId1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="3400" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="theme/theme1.xml"/></Relationships>`,
	);
	zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8"?>${DRAWINGS_HEADER_XML}`);
	zip.file(
		'word/_rels/header1.xml.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/image" Target="media/logo.png"/></Relationships>`,
	);
	zip.file('word/media/logo.png', LOGO_PNG);
	zip.file('word/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8"?>${THEME_XML}`);
	return zip.generateAsync({ type: 'uint8array' });
}
