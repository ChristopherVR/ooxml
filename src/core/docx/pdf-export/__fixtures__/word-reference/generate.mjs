// Writes the reference documents Word's Save as PDF is compared with (`../../word-reference.test.ts`):
// a letter, a multi-page table with a repeated header row, a header and footer with page numbers,
// footnotes and a landscape section, each in one of the free fonts the subsets in `../fonts` hold (Carlito, Liberation Sans and Serif,
// Noto Serif).
// Run with `bun src/core/docx/pdf-export/__fixtures__/word-reference/generate.mjs`; then measure
// them with `src/core/scripts/docx-word-pdf-reference.ps1`.
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

// Deterministic filler text: a fixed word list walked by a linear congruential generator.
const WORDS =
	'the of and a to in is you that it he was for on are as with his they at be this have from or one had by word but not what all were we when your can said there use an each which she do how their if will up other about out many then them these so some her would make like him into time has look two more write go see number no way could people my than first water been call who oil its now find long down day did get come made may part over new sound take only little work know place year live me back give most very after thing our just name good sentence man think say great where help through much before line right too mean old any same tell boy follow came want show also around form three small set put end does another well large must big even such because turn here why ask went men read need land different home us move try kind hand picture again change off play spell air away animal house point page letter mother answer found study still learn should America world'.split(
		' ',
	);
function filler(seed) {
	let state = seed >>> 0;
	return (count) => {
		const out = [];
		for (let index = 0; index < count; index++) {
			state = (Math.imul(state, 1103515245) + 12345) >>> 0;
			out.push(WORDS[(state >>> 8) % WORDS.length]);
		}
		const text = out.join(' ');
		return `${text[0].toUpperCase()}${text.slice(1)}.`;
	};
}

const esc = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const run = (text, props = '') =>
	`<w:r>${props && `<w:rPr>${props}</w:rPr>`}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const para = (content, props = '') => `<w:p>${props && `<w:pPr>${props}</w:pPr>`}${content}</w:p>`;
const field = (code) =>
	`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const LETTER = '<w:pgSz w:w="12240" w:h="15840"/>';
const LANDSCAPE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>';
const MARGINS =
	'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const sect = (size = LETTER, extra = '') => `<w:sectPr>${extra}${size}${MARGINS}</w:sectPr>`;

function styles(font, halfPoints, widowControl = true) {
	const fonts = `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`;
	return `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${fonts}<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${widowControl ? '<w:widowControl/>' : '<w:widowControl w:val="0"/>'}<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style><w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
}

async function write(name, { font, halfPoints, body, header, footer, footnotes, widowControl }) {
	const zip = new JSZip();
	const part = (path, xml) =>
		zip.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`, {
			date: new Date('2000-01-01T00:00:00Z'),
			createFolders: false,
		});
	const overrides = [
		['/word/document.xml', `${MAIN}.document.main+xml`],
		['/word/styles.xml', `${MAIN}.styles+xml`],
		['/word/settings.xml', `${MAIN}.settings+xml`],
		...(header ? [['/word/header1.xml', `${MAIN}.header+xml`]] : []),
		...(footer ? [['/word/footer1.xml', `${MAIN}.footer+xml`]] : []),
		...(footnotes ? [['/word/footnotes.xml', `${MAIN}.footnotes+xml`]] : []),
	];
	part(
		'[Content_Types].xml',
		`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([path, type]) => `<Override PartName="${path}" ContentType="${type}"/>`).join('')}</Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	const rels = [
		['rId1', 'styles', 'styles.xml'],
		['rId2', 'settings', 'settings.xml'],
		...(header ? [['rIdH', 'header', 'header1.xml']] : []),
		...(footer ? [['rIdF', 'footer', 'footer1.xml']] : []),
		...(footnotes ? [['rIdN', 'footnotes', 'footnotes.xml']] : []),
	];
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${PKG}">${rels.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`,
	);
	part('word/styles.xml', styles(font, halfPoints, widowControl));
	part(
		'word/settings.xml',
		`<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>${footnotes ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : ''}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
	);
	if (header) part('word/header1.xml', `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${header}</w:hdr>`);
	if (footer) part('word/footer1.xml', `<w:ftr xmlns:w="${W}" xmlns:r="${R}">${footer}</w:ftr>`);
	if (footnotes)
		part(
			'word/footnotes.xml',
			`<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${footnotes}</w:footnotes>`,
		);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`,
	);
	await writeFile(
		new URL(`./${name}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}

// 1. A one-page letter (Carlito 11 pt, Word's default spacing), with one justified paragraph.
{
	const text = filler(1);
	const lines = ['Jordan Example', '12 Sample Street', 'Exampletown 1234', '', 'November 27, 2026', ''];
	const body = [
		...lines.map((line) => para(line ? run(line) : '', '<w:spacing w:after="0"/>')),
		para(run('Dear Alex,')),
		para(run(text(70))),
		para(run(text(95)), '<w:jc w:val="both"/>'),
		para(run(text(55))),
		para(run('Yours sincerely,')),
		para(run('Jordan Example')),
	].join('');
	await write('letter', { font: 'Carlito', halfPoints: 22, body: body + sect() });
}

// 2. A table over three pages whose first row repeats as a header (Liberation Sans 10 pt).
{
	const text = filler(2);
	const cell = (width, content) =>
		`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
	const cellPara = (value, bold = false) =>
		para(run(value, bold ? '<w:b/>' : ''), '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>');
	const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
	const head = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(1200, cellPara('Item', true))}${cell(2400, cellPara('Name', true))}${cell(5760, cellPara('Description', true))}</w:tr>`;
	const rows = Array.from(
		{ length: 70 },
		(_, index) =>
			`<w:tr>${cell(1200, cellPara(String(index + 1)))}${cell(2400, cellPara(text(2 + (index % 3))))}${cell(5760, cellPara(text(8 + ((index * 7) % 25))))}</w:tr>`,
	).join('');
	const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="5760"/></w:tblGrid>${head}${rows}</w:tbl>`;
	const body = para(run('Inventory')) + table + para('') + sect();
	await write('table-header', { font: 'Liberation Sans', halfPoints: 20, body });
	// The same table with widow control off: Word splits the row that the control moves whole.
	await write('table-header-no-widow-control', {
		font: 'Liberation Sans',
		halfPoints: 20,
		body,
		widowControl: false,
	});
}

// 3. A header and a footer with "Page N of M" over two pages (Noto Serif 11 pt).
{
	const text = filler(3);
	const body =
		Array.from({ length: 18 }, (_, index) => para(run(text(40 + ((index * 13) % 50))))).join('') +
		sect(LETTER, '<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>');
	await write('header-footer', {
		font: 'Noto Serif',
		halfPoints: 22,
		body,
		header: para(run('Quarterly report') + run('\t') + run('\tDraft'), '<w:pStyle w:val="Header"/>'),
		footer: para(
			run('\tPage ') + field('PAGE') + run(' of ') + field('NUMPAGES'),
			'<w:pStyle w:val="Footer"/>',
		),
	});
}

// 4. Footnotes (Liberation Serif 12 pt) on two pages.
{
	const text = filler(4);
	let notes = '';
	const reference = (id) =>
		`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
	const paragraphs = Array.from({ length: 12 }, (_, index) => {
		const id = index + 1;
		notes += `<w:footnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${run(` ${text(10 + ((index * 11) % 30))}`)}</w:p></w:footnote>`;
		return para(run(text(45 + ((index * 17) % 40))) + reference(id) + run(` ${text(12)}`));
	}).join('');
	await write('footnotes', {
		font: 'Liberation Serif',
		halfPoints: 24,
		body: paragraphs + sect(),
		footnotes: notes,
	});
}

// 5. A portrait section, a landscape one and a portrait one again (Carlito 11 pt).
{
	const text = filler(5);
	const block = (count, seed) =>
		Array.from({ length: count }, (_, index) => para(run(text(35 + ((index + seed) * 19) % 60)))).join(
			'',
		);
	const breakTo = (size) => para('', `<w:sectPr>${size}${MARGINS}</w:sectPr>`);
	const body = block(9, 0) + breakTo(LETTER) + block(10, 3) + breakTo(LANDSCAPE) + block(6, 7) + sect();
	await write('landscape', { font: 'Carlito', halfPoints: 22, body });
}
