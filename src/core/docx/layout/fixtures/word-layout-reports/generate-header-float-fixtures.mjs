import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

// Synthetic fixtures for #401, #412, #416 and #408 (Arial 11, compatibility mode 15, A4).
// Regenerate with `bun generate-header-float-fixtures.mjs` from `src/core`.
const ns = {
	w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
};
const rootNs = Object.entries(ns)
	.map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
	.join(' ');
const office = ns.r;
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';
const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const margins =
	'w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"';

const para = (text, properties = '') =>
	`<w:p>${properties && `<w:pPr>${properties}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const pageField = (cached = '1') =>
	`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

/** A floating text box (`wp:anchor` with `wps:wsp`) holding `content`. */
function textBox({ cx, cy, h, v, wrap = '<wp:wrapNone/>', dist = '', effect = '', content }) {
	const attrs = `distT="0" distB="0" ${dist || 'distL="114300" distR="114300"'}`;
	return `<w:r><w:drawing><wp:anchor ${attrs} simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${h}${v}<wp:extent cx="${cx}" cy="${cy}"/>${effect || '<wp:effectExtent l="0" t="0" r="0" b="0"/>'}${wrap}<wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${ns.wps}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}

async function build(file, { header, footer, body, pgMar }) {
	const zip = new JSZip();
	const part = (name, xml) =>
		zip.file(name, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
	part(
		'[Content_Types].xml',
		`<Types xmlns="${types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${wml}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${wml}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${wml}.settings+xml"/>${header ? `<Override PartName="/word/header1.xml" ContentType="${wml}.header+xml"/>` : ''}${footer ? `<Override PartName="/word/footer1.xml" ContentType="${wml}.footer+xml"/>` : ''}</Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${office}/settings" Target="settings.xml"/>${header ? `<Relationship Id="rId3" Type="${office}/header" Target="header1.xml"/>` : ''}${footer ? `<Relationship Id="rId4" Type="${office}/footer" Target="footer1.xml"/>` : ''}</Relationships>`,
	);
	part(
		'word/styles.xml',
		`<w:styles xmlns:w="${ns.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
	);
	part(
		'word/settings.xml',
		`<w:settings xmlns:w="${ns.w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
	);
	if (header) part('word/header1.xml', `<w:hdr ${rootNs}>${header}</w:hdr>`);
	if (footer) part('word/footer1.xml', `<w:ftr ${rootNs}>${footer}</w:ftr>`);
	const refs = `${header ? '<w:headerReference w:type="default" r:id="rId3"/>' : ''}${footer ? '<w:footerReference w:type="default" r:id="rId4"/>' : ''}`;
	part(
		'word/document.xml',
		`<w:document ${rootNs}><w:body>${body}<w:sectPr>${refs}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar ${pgMar}/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(`./${file}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}

// #401: spacing after the last paragraph of a header and a footer.
for (const [name, after] of [
	['issue-401-header-footer-space-after', 200],
	['issue-401-header-footer-no-space-after', 0],
]) {
	const last = `<w:spacing w:after="${after}"/>`;
	await build(name, {
		header:
			[1, 2, 3, 4].map((n) => para(`Header line ${n}`)).join('') + para('Header line 5', last),
		footer: para('Footer line 1') + para('Footer line 2', last),
		body: para('First body line') + para('Second body line'),
		pgMar: margins,
	});
}

// #412: PAGE inside a text box in the footer, and the control with PAGE in the footer paragraph.
const pages = (n) =>
	Array.from({ length: n }, (_, i) =>
		para(`Body page ${i + 1}`, i ? '<w:pageBreakBefore/>' : ''),
	).join('');
await build('issue-412-page-field-in-footer-text-box', {
	footer: `<w:p><w:r><w:t>Footer</w:t></w:r>${textBox({
		cx: 914400,
		cy: 274320,
		h: '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>',
		v: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
		content: `<w:p>${pageField()}</w:p>`,
	})}</w:p>`,
	body: pages(3),
	pgMar: margins,
});
await build('issue-412-page-field-in-footer-paragraph', {
	footer: `<w:p><w:r><w:t xml:space="preserve">Footer </w:t></w:r>${pageField()}</w:p>`,
	body: pages(3),
	pgMar: margins,
});

// #416: distR and effectExtent of a square-wrapped text box at the left of the column.
const wrapped = (dist, effect, indent) => {
	const box = textBox({
		cx: 2009775,
		cy: 1371600,
		h: '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>',
		v: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
		wrap: '<wp:wrapSquare wrapText="bothSides"/>',
		dist,
		effect,
		content: para('Box'),
	});
	const properties = indent ? '<w:ind w:right="81"/>' : '';
	const text = 'Text beside the box ends with a short word at the edge so ';
	return (
		`<w:p>${properties && `<w:pPr>${properties}</w:pPr>`}${box}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` +
		Array.from({ length: 3 }, () => para(`${text}${text}`, properties)).join('')
	);
};
await build('issue-416-default-wrap-control', { body: wrapped('', '', false), pgMar: margins });
await build('issue-416-effect-extent-wrap', {
	body: wrapped('', '<wp:effectExtent l="0" t="0" r="9525" b="0"/>', true),
	pgMar: margins,
});
await build('issue-416-wrap-distance-36pt', {
	body: wrapped(
		'distL="114300" distR="457200"',
		'<wp:effectExtent l="0" t="0" r="38100" b="0"/>',
		false,
	),
	pgMar: margins,
});

// #408: a footer text box 7 % down the bottom margin (A4, bottom margin 1418).
const pctBox = (v) =>
	`<w:p>${textBox({
		cx: 368300,
		cy: 274320,
		h: '<wp:positionH relativeFrom="rightMargin"><wp:align>left</wp:align></wp:positionH>',
		v,
		content: para('8'),
	})}</w:p>`;
const pct =
	'<wp:positionV relativeFrom="bottomMargin"><wp14:pctPosVOffset>7000</wp14:pctPosVOffset></wp:positionV>';
const absolute =
	'<wp:positionV relativeFrom="page"><wp:posOffset>9854565</wp:posOffset></wp:positionV>';
const bottomMargins = margins.replace('w:bottom="1417"', 'w:bottom="1418"');
for (const [name, v] of [
	[
		'issue-408-pct-offset-alternate-content',
		`<mc:AlternateContent><mc:Choice Requires="wp14">${pct}</mc:Choice><mc:Fallback>${absolute}</mc:Fallback></mc:AlternateContent>`,
	],
	['issue-408-pct-offset-only', pct],
	['issue-408-absolute-offset-control', absolute],
])
	await build(name, { footer: pctBox(v), body: para('Body'), pgMar: bottomMargins });
