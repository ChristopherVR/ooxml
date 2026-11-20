import JSZip from 'jszip';

// Text boxes covering every `wps:bodyPr` layout setting the page engine lays out, each floating in
// front of the text at a fixed page position (EMU from the page's top-left) so its lines can be
// compared with Word's PDF of the same file: a wrapped box that overflows, two columns, `vert`,
// `vert270`, `spAutoFit` with a stored height too small for its text, `wrap="none"`, `normAutofit`,
// mixed formatting anchored in the middle, and a linked pair (`wps:txbx id` and `wps:linkedTxbx`).

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const NAMESPACES = [
	`xmlns:w="${W}"`,
	`xmlns:r="${R}"`,
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(' ');
const INCH = 914400;

/** Numbered words, so a line break or a linked box's split can be located in Word's PDF. */
export const words = (label: string, count: number): string =>
	[label, ...Array.from({ length: count }, (_, index) => `word${index + 1}`)].join(' ');

export interface FixtureBox {
	id: number;
	name: string;
	/** Inches from the page's top-left. */
	x: number;
	y: number;
	width: number;
	height: number;
	/** Extra `wps:bodyPr` attributes (`vert`, `numCol`, `wrap`, `anchor`) and its autofit child. */
	bodyPr?: Record<string, string>;
	autofit?: 'noAutofit' | 'spAutoFit' | 'normAutofit';
	/** `w:txbxContent` paragraphs, or the chain this box continues. */
	content: string | { linked: number; seq: number };
	/** The chain id this box starts (`wps:txbx id`). */
	chain?: number;
	/** Outline width in EMU (0.75 pt when absent); 0 draws no outline. */
	outline?: number;
	/** The shape's preset (`rect` when absent). */
	preset?: string;
}

const paragraph = (text: string, pPr = '', rPr = '') =>
	`<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

export const BOXES: FixtureBox[] = [
	{
		id: 1,
		name: 'Wrapped',
		x: 1,
		y: 1.5,
		width: 2,
		height: 1,
		content: paragraph(words('Wrapped', 40)),
	},
	{
		id: 2,
		name: 'Columns',
		x: 4,
		y: 1.5,
		width: 3,
		height: 1.2,
		bodyPr: { numCol: '2', spcCol: String(0.2 * INCH) },
		content: paragraph(words('Columns', 40)),
	},
	{
		id: 3,
		name: 'Vertical',
		x: 1,
		y: 3,
		width: 0.9,
		height: 2.4,
		bodyPr: { vert: 'vert' },
		content: paragraph(words('Vertical', 12)),
	},
	{
		id: 4,
		name: 'Vertical 270',
		x: 2.5,
		y: 3,
		width: 0.9,
		height: 2.4,
		bodyPr: { vert: 'vert270' },
		content: paragraph(words('Upward', 12)),
	},
	{
		id: 5,
		name: 'Autofit',
		x: 4,
		y: 3,
		width: 2.5,
		height: 0.4,
		autofit: 'spAutoFit',
		content: paragraph(words('Autofit', 24)),
	},
	{
		id: 6,
		name: 'No wrap',
		x: 4,
		y: 4.6,
		width: 1.5,
		height: 0.4,
		bodyPr: { wrap: 'none' },
		content: paragraph('Unwrapped line that runs past the box'),
	},
	{
		id: 7,
		name: 'Linked first',
		x: 1,
		y: 6,
		width: 2,
		height: 0.8,
		chain: 1,
		content: paragraph(words('Linked', 45)),
	},
	{
		id: 8,
		name: 'Linked second',
		x: 4,
		y: 6,
		width: 2.5,
		height: 1.2,
		content: { linked: 1, seq: 1 },
	},
	{
		id: 9,
		name: 'Shrink',
		x: 1,
		y: 8,
		width: 2,
		height: 0.6,
		autofit: 'normAutofit',
		content: paragraph(words('Shrink', 24)),
	},
	{
		id: 10,
		name: 'Formatted',
		x: 4,
		y: 8,
		width: 2.5,
		height: 1.2,
		bodyPr: { anchor: 'ctr' },
		content:
			paragraph('Bold heading', '<w:jc w:val="right"/>', '<w:b/><w:sz w:val="28"/>') +
			paragraph(
				'Small italic text under a border',
				'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="1F3864"/></w:pBdr>',
				'<w:i/><w:sz w:val="16"/>',
			),
	},
];

function boxXml(box: FixtureBox): string {
	const emu = (inches: number) => Math.round(inches * INCH);
	const cx = emu(box.width);
	const cy = emu(box.height);
	const body = {
		rot: '0',
		vert: 'horz',
		wrap: 'square',
		lIns: '91440',
		tIns: '45720',
		rIns: '91440',
		bIns: '45720',
		anchor: 't',
		anchorCtr: '0',
		...box.bodyPr,
	};
	const attributes = Object.entries(body)
		.map(([name, value]) => `${name}="${value}"`)
		.join(' ');
	const text =
		typeof box.content === 'string'
			? `<wps:txbx${box.chain ? ` id="${box.chain}"` : ''}><w:txbxContent>${box.content}</w:txbxContent></wps:txbx>`
			: `<wps:linkedTxbx id="${box.content.linked}" seq="${box.content.seq}"/>`;
	return (
		`<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + box.id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
		`<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(box.x)}</wp:posOffset></wp:positionH>` +
		`<wp:positionV relativeFrom="page"><wp:posOffset>${emu(box.y)}</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
		`<wp:docPr id="${box.id}" name="${box.name}"/><wp:cNvGraphicFramePr/>` +
		`<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
		`<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${box.preset ?? 'rect'}"><a:avLst/></a:prstGeom>` +
		'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
		(box.outline === 0
			? '<a:ln><a:noFill/></a:ln>'
			: `<a:ln w="${box.outline ?? 9525}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`) +
		'</wps:spPr>' +
		`${text}<wps:bodyPr ${attributes}><a:${box.autofit ?? 'noAutofit'}/></wps:bodyPr>` +
		'</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
	);
}

/** Arial 11 pt, single spacing and no paragraph spacing, so lines compare with Word's directly. */
const STYLES_XML = `<?xml version="1.0" encoding="UTF-8"?><w:styles ${NAMESPACES}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

/** Word 2013 and later layout (compatibility mode 15), as Word saves new documents. */
const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8"?><w:settings ${NAMESPACES}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

/** A document holding `boxes` in its first paragraph. */
export const textBoxDocumentXml = (boxes: FixtureBox[]): string =>
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>` +
	`<w:p>${boxes.map(boxXml).join('')}<w:r><w:t>Text box layout: each box is placed on the page in front of the text.</w:t></w:r></w:p>` +
	'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>';

/** The fixture package: the document, its styles and nothing else. */
export async function textBoxLayoutFixture(boxes: FixtureBox[] = BOXES): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file('word/document.xml', textBoxDocumentXml(boxes));
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/settings" Target="settings.xml"/></Relationships>`,
	);
	zip.file('word/styles.xml', STYLES_XML);
	zip.file('word/settings.xml', SETTINGS_XML);
	return zip.generateAsync({ type: 'uint8array' });
}
