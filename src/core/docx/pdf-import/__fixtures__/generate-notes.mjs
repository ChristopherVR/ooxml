// The Word sources of the notes, links and rotated text fixtures (run through
// generate.mjs). Word 16 exports them to PDF (export-word.ps1), with heading bookmarks and tags.
import { writeFile, mkdir } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
const DATE = new Date('2026-11-24T00:00:00Z');
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const esc = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const run = (text, props = '') =>
	`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const para = (runs, props = '') => `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}${runs}</w:p>`;
const heading = (level, text, bookmark) =>
	para(
		(bookmark ? `<w:bookmarkStart w:id="${bookmark.id}" w:name="${bookmark.name}"/>` : '') +
			run(text) +
			(bookmark ? `<w:bookmarkEnd w:id="${bookmark.id}"/>` : ''),
		`<w:pStyle w:val="Heading${level}"/>`,
	);
const note = (id) =>
	`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;

const TEXT = [
	'Ce guide présente une démarche générique, illustrée par des exemples sans lien avec une organisation réelle.',
	'Chaque section reprend les mêmes principes et les applique à un cas différent, pas à pas.',
	'Les paragraphes sont assez longs pour occuper plusieurs lignes de la page imprimée.',
];
const body = (count) => Array.from({ length: count }, (_, index) => TEXT[index % TEXT.length]).join(' ');

const STYLES = `${DECL}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9062"/></w:tabs><w:spacing w:after="100"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9062"/></w:tabs><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`;

const SETTINGS = `${DECL}<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const SECTION = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;

async function save(name, { document, parts = {}, rels = [], overrides = [] }) {
	const zip = new JSZip();
	const add = (path, data) => zip.file(path, data, { date: DATE, createFolders: false });
	add(
		'[Content_Types].xml',
		`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${WML}.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${overrides.join('')}</Types>`,
	);
	add(
		'_rels/.rels',
		`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
	);
	add(
		'docProps/core.xml',
		`${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(name)}</dc:title><dc:creator>Exemple</dc:creator></cp:coreProperties>`,
	);
	add(
		'word/_rels/document.xml.rels',
		`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/>${rels.join('')}</Relationships>`,
	);
	add('word/styles.xml', STYLES);
	add('word/settings.xml', SETTINGS);
	add('word/document.xml', `${DECL}<w:document ${NS}><w:body>${document}${SECTION}</w:body></w:document>`);
	for (const [path, data] of Object.entries(parts)) add(path, data);
	await mkdir(new URL('./source/', import.meta.url), { recursive: true });
	await writeFile(
		new URL(`./source/${name}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
	);
}

// notes.docx: headings with a table of contents, footnotes on two pages and an internal link.
const footnotes = [
	'Une première note de bas de page, courte.',
	'Une deuxième note, plus longue, qui explique un point de détail du paragraphe sans encombrer le texte principal.',
	'La note de la deuxième page.',
];
const toc =
	para(
		'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
			run('Table des matières à mettre à jour.'),
		'<w:pStyle w:val="TOC1"/>',
	) + para('<w:r><w:fldChar w:fldCharType="end"/></w:r>');
const notesDocument = [
	para(run('Guide pratique', '<w:b/><w:sz w:val="44"/>'), '<w:jc w:val="center"/>'),
	toc,
	heading(1, 'Introduction'),
	para(run(body(3)) + note(1) + run(' ' + body(2)), '<w:jc w:val="both"/>'),
	heading(2, 'Contexte'),
	para(run(body(4)) + note(2) + run(' La suite du paragraphe précise le contexte.'), '<w:jc w:val="both"/>'),
	para(
		run('Pour les détails, ') +
			`<w:hyperlink w:anchor="annexe" w:history="1">${run('voir l’annexe', '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>` +
			run(' en fin de document.'),
	),
	heading(1, 'Méthode'),
	para(run(body(5)), '<w:jc w:val="both"/>'),
	para('<w:r><w:br w:type="page"/></w:r>'),
	heading(2, 'Étapes'),
	para(run(body(3)) + note(3) + run(' ' + body(1)), '<w:jc w:val="both"/>'),
	heading(1, 'Annexe', { id: 10, name: 'annexe' }),
	para(run(body(2)), '<w:jc w:val="both"/>'),
].join('');
const footnotesXml =
	`${DECL}<w:footnotes ${NS}>` +
	'<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
	'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
	footnotes
		.map(
			(text, index) =>
				`<w:footnote w:id="${index + 1}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${run(` ${text}`)}</w:p></w:footnote>`,
		)
		.join('') +
	'</w:footnotes>';
await save('notes', {
	document: notesDocument,
	parts: { 'word/footnotes.xml': footnotesXml },
	rels: [`<Relationship Id="rIdFootnotes" Type="${R}/footnotes" Target="footnotes.xml"/>`],
	overrides: [`<Override PartName="/word/footnotes.xml" ContentType="${WML}.footnotes+xml"/>`],
});

// rotation.docx: body text with two text boxes whose text runs bottom to top and top to bottom.
const textBox = (id, text, vert, x) =>
	`<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>3600000</wp:posOffset></wp:positionV><wp:extent cx="457200" cy="2160000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Zone de texte ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="2160000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${para(run(text, '<w:sz w:val="24"/>'), '<w:spacing w:after="0"/>')}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="${vert}" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>`;
await save('rotation', {
	document: [
		para(run('Texte tourné', '<w:b/><w:sz w:val="32"/>')),
		para(
			textBox(1, 'Écrit de bas en haut', 'vert270', 180000) +
				textBox(2, 'Écrit de haut en bas', 'vert', 6840000) +
				run(body(6)),
			'<w:jc w:val="both"/>',
		),
		para(run(body(3)), '<w:jc w:val="both"/>'),
	].join(''),
});
