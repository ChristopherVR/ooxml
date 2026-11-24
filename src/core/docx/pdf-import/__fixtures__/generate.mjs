// Writes the Word source documents of the PDF import fixtures. Word 16 exports them to PDF
// (export-word.ps1) and pypdf encrypts copies (encrypt.py); the PDFs are what the tests read.
// Run from `src/core`: bun docx/pdf-import/__fixtures__/generate.mjs
import { writeFile, mkdir } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const DATE = new Date('2026-11-24T00:00:00Z');

const esc = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const run = (text, props = '') => `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const para = (runs, props = '') => `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}${runs}</w:p>`;
const field = (instr, result) =>
	`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const SENTENCES = [
	'Le présent rapport décrit l’activité de l’exercice et les perspectives de la société.',
	'Les chiffres présentés sont des exemples génériques sans lien avec une entreprise réelle.',
	'Chaque paragraphe est assez long pour revenir à la ligne plusieurs fois dans la page.',
	'La mise en forme utilise des polices courantes, une couleur, du gras et de l’italique.',
];
const body = (count) => Array.from({ length: count }, (_, index) => SENTENCES[index % SENTENCES.length]).join(' ');

/** A 40 x 24 PNG: a blue band over a light background. */
function png() {
	const width = 40;
	const height = 24;
	const crcTable = Array.from({ length: 256 }, (_, n) => {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		return c >>> 0;
	});
	const crc = (bytes) => {
		let c = 0xffffffff;
		for (const b of bytes) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
		return (c ^ 0xffffffff) >>> 0;
	};
	const chunk = (type, data) => {
		const out = Buffer.alloc(12 + data.length);
		out.writeUInt32BE(data.length, 0);
		out.write(type, 4, 'latin1');
		data.copy(out, 8);
		out.writeUInt32BE(crc(out.subarray(4, 8 + data.length)), 8 + data.length);
		return out;
	};
	const header = Buffer.alloc(13);
	header.writeUInt32BE(width, 0);
	header.writeUInt32BE(height, 4);
	header.set([8, 2, 0, 0, 0], 8);
	const rows = [];
	for (let y = 0; y < height; y++) {
		rows.push(0);
		for (let x = 0; x < width; x++) rows.push(...(y > 8 && y < 16 ? [30, 80, 160] : [230, 236, 245]));
	}
	return Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', deflateSync(Buffer.from(rows))),
		chunk('IEND', Buffer.alloc(0)),
	]);
}

const picture = `<w:r><w:drawing><wp:inline><wp:extent cx="2286000" cy="1371600"/><wp:docPr id="1" name="Illustration"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="band.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const content = [
	para(run('Rapport annuel', '<w:b/><w:sz w:val="40"/>'), '<w:pStyle w:val="Heading1"/>'),
	para(run(body(6)), '<w:jc w:val="both"/><w:spacing w:after="160"/>'),
	para(
		run('Un passage en ') + run('gras', '<w:b/>') + run(', un en ') + run('italique', '<w:i/>') + run(', un ') + run('souligné', '<w:u w:val="single"/>') + run(', un ') + run('rouge', '<w:color w:val="C00000"/>') + run(' et un exposant m') + run('2', '<w:vertAlign w:val="superscript"/>') + run('.'),
		'<w:spacing w:after="160"/>',
	),
	para(run('Activité', '<w:b/><w:sz w:val="28"/>'), '<w:pStyle w:val="Heading2"/>'),
	para(run(body(10)), '<w:jc w:val="both"/><w:spacing w:after="160"/>'),
	para(picture, '<w:jc w:val="center"/><w:spacing w:after="160"/>'),
	para(run('Texte centré sous l’illustration', '<w:i/>'), '<w:jc w:val="center"/><w:spacing w:after="160"/>'),
	...Array.from({ length: 8 }, (_, index) => para(run(`${index + 1}. ${body(4)}`), '<w:jc w:val="both"/><w:spacing w:after="160"/>')),
	para(run('Perspectives', '<w:b/><w:sz w:val="28"/>'), '<w:pStyle w:val="Heading2"/>'),
	para(run(body(12)), '<w:jc w:val="both"/><w:spacing w:after="160"/>'),
	para(run('Fin du rapport.')),
].join('');

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="240"/><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>`;

async function save(name, documentXml, parts) {
	const zip = new JSZip();
	const add = (path, data) => zip.file(path, data, { date: DATE, createFolders: false });
	add('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
	add('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
	add('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Rapport annuel</dc:title><dc:creator>Exemple</dc:creator></cp:coreProperties>`);
	add('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="${R}/footer" Target="footer1.xml"/><Relationship Id="rIdImage" Type="${R}/image" Target="media/band.png"/></Relationships>`);
	add('word/styles.xml', styles);
	add('word/media/band.png', png());
	add('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}>${para(run('Société Exemple – Rapport annuel 2026'), '<w:jc w:val="right"/>')}</w:hdr>`);
	add('word/footer1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}>${para(run('Page ') + field('PAGE', '1') + run(' sur ') + field('NUMPAGES', '3'), '<w:jc w:val="center"/>')}</w:ftr>`);
	add('word/document.xml', documentXml);
	for (const [path, data] of Object.entries(parts)) add(path, data);
	await mkdir(new URL('./source/', import.meta.url), { recursive: true });
	await writeFile(new URL(`./source/${name}.docx`, import.meta.url), await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}

await save(
	'rapport',
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${content}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	{},
);
