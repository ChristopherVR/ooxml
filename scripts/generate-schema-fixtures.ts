// Generates representative .docx, .xlsx, .pptx and .vsdx packages from the core writers into an
// output directory so CI can validate them with Microsoft's OpenXmlValidator and the Visio
// structural package validator.
// Usage: bun scripts/generate-schema-fixtures.ts <output-dir>
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';

import {
	buildTableOfContents,
	createDocument,
	ensureListDefinition,
	loadDocx,
	saveDocx,
	tocBookmarks,
	tocEntries,
	withTocBookmarks,
	type DocumentModel,
	type Paragraph,
} from '../src/core/docx/index';
import { convertPdf } from '../src/core/docx/pdf-import/import';
import { eighthPoints, halfPoints, signedTwips, twips } from '../src/core/docx/units';
import { createPresentationFromOutline } from '../src/core/pptx/core/utils/suite-text';
import { mergePptxSnapshots } from '../src/core/pptx/merge-snapshots';
import { createVsdx } from '../src/core/visio/create-document';
import { editVsdx } from '../src/core/visio/edit';
import { createWorkbook, createEditSession, saveXlsx } from '../src/core/xlsx/index';
import { NS } from '../src/core/xml/namespaces';
import { buildXml, elements, parseXml } from '../src/core/xml/xml';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PNG = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
		'base64',
	),
);
const SVG = new TextEncoder().encode(
	'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>',
);

const paragraphs = (model: DocumentModel) =>
	model.blocks.filter((block): block is Paragraph => block.type === 'paragraph');

async function packageWith(files: Record<string, string>): Promise<Uint8Array> {
	const zip = new JSZip();
	if (!files['_rels/.rels']) {
		zip.file(
			'_rels/.rels',
			rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`),
		);
	}
	if (!files['[Content_Types].xml']) {
		const overrides = Object.keys(files)
			.filter((name) => name.endsWith('.xml'))
			.map((name) => {
				const type =
					name === 'word/document.xml'
						? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
						: name === 'word/endnotes.xml'
							? 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'
							: /^word\/header\d*\.xml$/.test(name)
								? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
								: /^word\/footer\d*\.xml$/.test(name)
									? 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
									: 'application/xml';
				return `<Override PartName="/${name}" ContentType="${type}"/>`;
			})
			.join('');
		zip.file(
			'[Content_Types].xml',
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
		);
	}
	for (const [name, contents] of Object.entries(files)) zip.file(name, contents);
	return zip.generateAsync({ type: 'uint8array' });
}

const rels = (entries: string) =>
	`<?xml version="1.0"?><Relationships xmlns="${REL_PKG}">${entries}</Relationships>`;
const wDocument = (body: string) =>
	`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`;

function docxKitchenSink(): DocumentModel {
	const model = createDocument();
	const rev = { author: 'Ada', date: '2024-01-01T00:00:00Z' };
	model.blocks = [
		{ type: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: 'Heading' }], keepNext: true },
		{
			type: 'paragraph',
			id: 'p1',
			align: 'justify',
			spacingBeforeTwips: twips(120),
			lineSpacingTwips: signedTwips(360),
			lineSpacingRule: 'auto',
			indentLeftTwips: signedTwips(720),
			hangingTwips: twips(360),
			tabStops: [{ posTwips: signedTwips(4680), align: 'right', leader: 'dot' }],
			pageBreakBefore: true,
			direction: 'rtl',
			runs: [
				{
					text: 'Bold',
					bold: true,
					italic: false,
					underline: true,
					strike: false,
					color: '#FF0000',
					fontSize: 14,
					fontFamily: 'Arial',
					highlight: 'yellow',
					verticalAlign: 'superscript',
					language: 'en-US',
					rtl: true,
					caps: true,
					smallCaps: false,
					vanish: false,
					characterSpacingTwips: signedTwips(20),
					textScalePercent: 125,
					ligatures: 'standardContextual',
					kerningHalfPoints: halfPoints(24),
					positionHalfPoints: halfPoints(-6),
					shadingFill: '#FFFF00',
					underlineStyle: 'double',
					underlineColor: '#00FF00',
					doubleStrike: false,
					style: 'Hyperlink',
				},
				{ text: 'baseline', verticalAlign: 'baseline' },
				{ text: 'link', link: { anchor: '_Toc100000000' } },
				{ text: '', break: 'page' },
				{ text: '\tafter tab\n' },
				{ text: 'inserted', revision: { ...rev, kind: 'insert', id: '5' } },
				{ text: 'deleted', revision: { ...rev, kind: 'delete', id: '6' } },
				{ text: 'moved', revision: { ...rev, kind: 'moveFrom', id: '7', move: { name: 'move1' } } },
				{ text: 'moved', revision: { ...rev, kind: 'moveTo', id: '8', move: { name: 'move1' } } },
				{ text: 'commented', commentIds: ['c1'] },
				{ text: '', noteReference: { kind: 'footnote', id: '1' } },
				{ text: '', fieldChar: 'begin' },
				{ text: '', fieldCode: ' PAGE ' },
				{ text: '', fieldChar: 'separate' },
				{ text: '1', field: { instr: 'PAGE' } },
				{ text: '', fieldChar: 'end' },
				{ text: 'Ann', field: { instr: 'AUTHOR', simple: true } },
			],
			markRevision: { ...rev, kind: 'insert', id: '9' },
		},
		{
			type: 'table',
			id: 't1',
			rows: [
				[
					{
						paragraphs: [{ type: 'paragraph', id: 'c1p', runs: [{ text: 'A' }] }],
						gridSpan: 1,
						verticalAlign: 'center',
						shadingFill: 'D9D9D9',
						widthTwips: twips(2000),
					},
					{ paragraphs: [{ type: 'paragraph', id: 'c2p', runs: [{ text: 'B' }] }] },
				],
			],
		},
		{ type: 'paragraph', id: 'end', runs: [{ text: 'End' }] },
	];
	model.comments = [{ id: 'c1', author: 'Ada', text: 'Note this', resolved: true }];
	model.footnotes = [
		{
			id: '1',
			blocks: [
				{
					type: 'paragraph',
					id: 'n1',
					style: 'FootnoteText',
					runs: [
						{ text: '', noteMark: 'footnote', style: 'FootnoteReference' },
						{ text: ' Source' },
					],
				},
			],
		},
	];
	return model;
}

function sinkWithToc(): DocumentModel {
	const model = docxKitchenSink();
	const { bookmarks, added } = tocBookmarks(model, tocEntries(model));
	const marked = withTocBookmarks(model, added);
	let next = 0;
	marked.blocks.unshift(
		...buildTableOfContents(marked, {
			newId: () => `toc${++next}`,
			bookmarks,
			pageNumbers: new Map([['h1', '1']]),
		}),
	);
	return marked;
}

async function buildDocxFixtures(): Promise<Record<string, Uint8Array>> {
	const fixtures: Record<string, Uint8Array> = {};

	// 1. Kitchen sink + TOC
	const sinkBytes = await saveDocx(sinkWithToc());
	fixtures['docx-kitchen-sink.docx'] = sinkBytes;

	// 2. Loaded, edited, with new numbering list
	{
		const loaded = await loadDocx(sinkBytes);
		const model = structuredClone(loaded.model);
		const { catalog, numId } = ensureListDefinition(model.numberingCatalog, 'decimal');
		model.numberingCatalog = catalog;
		const plain = paragraphs(model).filter((p) => !p.style);
		plain[0]!.runs[0]!.text = 'Edited text';
		plain.at(-1)!.numbering = { numId, level: 0 };
		fixtures['docx-edited-list.docx'] = await loaded.save(model);
	}

	// 3. Headers and footers shared by two sections, edited
	{
		const refs = `<w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>`;
		const bytes = await packageWith({
			'word/document.xml': wDocument(
				`<w:p><w:pPr><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
			),
			'word/_rels/document.xml.rels': rels(
				`<Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${R}/footer" Target="footer1.xml"/>`,
			),
			'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Draft</w:t></w:r></w:p></w:hdr>`,
			'word/footer1.xml': `<w:ftr xmlns:w="${W}"><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
		});
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		for (const section of model.sections ?? []) {
			const first = section.headers?.default?.blocks[0];
			if (first?.type === 'paragraph' && first.runs[0]) first.runs[0].text = 'Final';
		}
		fixtures['docx-headers-footers.docx'] = await loaded.save(model);
	}

	// 4. Endnotes edited + added
	{
		const bytes = await packageWith({
			'word/document.xml': wDocument(
				`<w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr/>`,
			),
			'word/endnotes.xml': `<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Source.</w:t></w:r></w:p></w:endnote></w:endnotes>`,
		});
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		const firstNoteBlock = model.endnotes?.[0]?.blocks[0];
		if (firstNoteBlock?.type === 'paragraph' && firstNoteBlock.runs[1])
			firstNoteBlock.runs[1].text = ' Edited source.';
		model.endnotes!.push({
			id: '3',
			blocks: [
				{
					type: 'paragraph',
					id: 'en3',
					runs: [{ text: '', noteMark: 'endnote' }, { text: ' Second.' }],
				},
			],
		});
		paragraphs(model)[0]!.runs.push({ text: '', noteReference: { kind: 'endnote', id: '3' } });
		fixtures['docx-endnotes.docx'] = await loaded.save(model);
	}

	// 5. PNG + resized SVG picture
	{
		const first = createDocument();
		first.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [
				{
					text: '',
					image: {
						relId: '',
						partName: 'word/media/p.png',
						svgPartName: 'word/media/p.svg',
						contentType: 'image/png',
						widthPx: 40,
						heightPx: 40,
					},
				},
			],
		};
		const bytes = await saveDocx(
			first,
			new Map([
				['word/media/p.png', { bytes: PNG, contentType: 'image/png' }],
				['word/media/p.svg', { bytes: SVG, contentType: 'image/svg+xml' }],
			]),
		);
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		const paragraph = paragraphs(model)[0]!;
		const image = paragraph.runs[0]!.image!;
		image.widthPx = 80;
		image.heightPx = 20;
		paragraph.runs.push({
			text: '',
			image: {
				relId: '',
				partName: 'word/media/n.png',
				contentType: 'image/png',
				widthPx: 16,
				heightPx: 16,
				altText: 'New',
			},
		});
		fixtures['docx-pictures-svg.docx'] = await loaded.save(
			model,
			new Map([['word/media/n.png', { bytes: PNG, contentType: 'image/png' }]]),
		);
	}

	// 6. Merged, shaded, bordered table
	{
		const model = createDocument();
		const cellParagraph = (id: string) => ({
			type: 'paragraph' as const,
			id,
			runs: [{ text: id }],
		});
		model.blocks.push({
			type: 'table',
			id: 'new-table',
			widthTwips: twips(9000),
			alignment: 'right',
			indentTwips: signedTwips(100),
			cellMargins: { left: twips(108), right: twips(108) },
			borders: {
				top: { style: 'single', sizeEighthPoints: eighthPoints(4), themeColor: 'accent1' },
			},
			rowProperties: [
				{ heightTwips: twips(500), heightRule: 'exact', cantSplit: true, header: true },
			],
			rows: [
				[
					{
						paragraphs: [cellParagraph('m1')],
						gridSpan: 2,
						verticalMerge: 'restart',
						widthTwips: twips(6000),
						shadingFill: '#D9D9D9',
						shadingThemeFill: { token: 'accent2', tint: 0.5 },
						verticalAlign: 'center',
						borders: {
							bottom: { style: 'double', sizeEighthPoints: eighthPoints(6), color: '#112233' },
						},
						margins: { top: twips(10), left: twips(20), bottom: twips(30), right: twips(40) },
					},
					{ paragraphs: [cellParagraph('m2')], widthTwips: twips(3000) },
				],
				[
					{ paragraphs: [cellParagraph('m3')], verticalMerge: 'continue', gridSpan: 2 },
					{ paragraphs: [cellParagraph('m4')] },
				],
				[{ paragraphs: [cellParagraph('m5')] }, { paragraphs: [cellParagraph('m6')] }],
			],
		});
		fixtures['docx-table-merged.docx'] = await saveDocx(model);
	}

	// 7. Section columns + header reference
	{
		const bytes = await packageWith({
			'word/document.xml': wDocument(
				`<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:cols w:space="720"/><w:titlePg/></w:sectPr>`,
			),
			'word/_rels/document.xml.rels': rels(
				`<Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/>`,
			),
			'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
		});
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		const last = model.sections![0]!;
		model.sections = [
			{ ...structuredClone(last), endsAtBlockId: model.blocks[0]!.id },
			{
				...last,
				type: 'continuous',
				columns: { count: 2, spacingTwips: twips(720), equalWidth: true },
			},
		];
		fixtures['docx-section-columns.docx'] = await loaded.save(model);
	}

	// 8. PDF-imported vector drawings (exercises Word 2010+ wps/wpg namespaces)
	{
		const pdf = new Uint8Array(
			readFileSync(
				new URL('../src/core/docx/pdf-import/__fixtures__/dessins.pdf', import.meta.url),
			),
		);
		const { docx } = await convertPdf(pdf);
		fixtures['docx-pdf-drawings.docx'] = docx;
	}

	return fixtures;
}

async function buildXlsxFixtures(): Promise<Record<string, Uint8Array>> {
	const wb = createWorkbook({ sheets: ['Summary', 'Data'] });
	const session = createEditSession(wb);
	session.setCellValue(0, 0, 0, 'Quarter');
	session.setCellValue(0, 0, 1, 'Revenue');
	session.setCellValue(0, 1, 0, 'Q1');
	session.setCellValue(0, 1, 1, 1250);
	session.setCellValue(0, 2, 0, 'Q2');
	session.setCellValue(0, 2, 1, 1890);
	session.setCellInput(0, 3, 0, 'Total');
	session.setCellInput(0, 3, 1, '=SUM(B2:B3)');
	session.setCellValue(1, 0, 0, 'Notes');
	session.setCellValue(1, 1, 0, 'Generated for OpenXmlValidator');
	return {
		'xlsx-workbook-basic.xlsx': await saveXlsx(wb),
	};
}

async function buildPptxFixtures(): Promise<Record<string, Uint8Array>> {
	const outline = await createPresentationFromOutline(
		[
			{ title: 'Quarterly Review' },
			{
				title: 'Highlights',
				bullets: ['Shipped schema validation', 'Reduced maintenance burden'],
				notes: 'Speaker notes for slide 2.',
			},
		],
		{ title: 'Schema Validation Fixture' },
	);

	// Also exercise mergePptxSnapshots on a real deck
	const baseFixture = readFileSync(
		new URL('../src/core/pptx/__tests__/fixtures/slide-size-centering.pptx', import.meta.url),
	);
	const slidePart = 'ppt/slides/slide1.xml';
	async function deckWithSpPr(children: string): Promise<Uint8Array> {
		const zip = await JSZip.loadAsync(baseFixture);
		const doc = parseXml(await zip.file(slidePart)!.async('string'));
		const shape = doc.getElementsByTagNameNS(NS.p, 'sp')[0]!;
		const spPr = elements(shape).find((child) => child.localName === 'spPr')!;
		for (const child of elements(spPr)) {
			if (!['xfrm', 'prstGeom', 'custGeom'].includes(child.localName)) spPr.removeChild(child);
		}
		const fragment = parseXml(`<root xmlns:a="${NS.a}">${children}</root>`);
		for (const child of elements(fragment.documentElement))
			spPr.appendChild(doc.importNode(child, true));
		zip.file(slidePart, buildXml(doc));
		return zip.generateAsync({ type: 'uint8array' });
	}
	const base = await deckWithSpPr('');
	const local = await deckWithSpPr('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>');
	const remote = await deckWithSpPr(
		'<a:ln w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>',
	);
	const merged = await mergePptxSnapshots(base, local, remote);
	if (merged.status !== 'merged')
		throw new Error(`mergePptxSnapshots failed: ${merged.issues.join('; ')}`);

	return {
		'pptx-outline.pptx': outline,
		'pptx-merged-snapshots.pptx': merged.bytes,
	};
}

async function buildVisioFixtures(): Promise<Record<string, Uint8Array>> {
	const blank = await createVsdx();
	const landscape = await createVsdx({ width: 11, height: 8.5 });
	const drawn = await editVsdx(landscape, [
		{
			type: 'create-rectangle',
			pageId: '0',
			shapeId: '1',
			x: 2,
			y: 3,
			width: 2,
			height: 1,
			text: 'Process Step',
		},
	]);
	const formatted = await editVsdx(drawn.bytes, [
		{ type: 'format-text', pageId: '0', shapeId: '1', bold: true },
	]);
	return {
		'visio-blank.vsdx': blank,
		'visio-shapes-edited.vsdx': formatted.bytes,
	};
}

async function main() {
	const outDir = resolve(process.argv[2] ?? 'tmp-schema-fixtures');
	mkdirSync(outDir, { recursive: true });
	const all = {
		...(await buildDocxFixtures()),
		...(await buildXlsxFixtures()),
		...(await buildPptxFixtures()),
		...(await buildVisioFixtures()),
	};
	for (const [name, bytes] of Object.entries(all)) {
		writeFileSync(join(outDir, name), bytes);
	}
	console.log(`Generated ${Object.keys(all).length} schema validation fixtures in ${outDir}`);
}

main().catch((error) => {
	console.error(error);
	process.exitCode = 1;
});
