import { describe, expect, it } from 'vitest';
import type { DocumentModel, Paragraph, ParagraphFrame } from '../index';
import {
	parseParagraphStyleCatalog,
	resolveParagraphFormatting,
	signedTwips,
	twips,
} from '../index';
import { parseDirectParagraphProperties } from '../paragraph-properties';
import { first, parseXml } from '../xml';
import { layoutDocument, layoutDocumentModel } from './layout';
import type { LayoutBlock, LayoutParagraph, LayoutTextFrame } from './input';
import type { TextMeasurer } from './measure';
import type { LayoutParagraphBox } from './result';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };

const page = {
	widthPx: 800,
	heightPx: 1000,
	marginTopPx: 100,
	marginRightPx: 100,
	marginBottomPx: 100,
	marginLeftPx: 100,
};

const words = Array.from({ length: 60 }, () => 'word').join(' ');

function pPrXml(inner: string) {
	const doc = parseXml(`<w:pPr xmlns:w="${W_NS}">${inner}</w:pPr>`);
	return doc.documentElement;
}

function paragraph(
	id: string,
	text = 'text',
	extra: Partial<LayoutParagraph> = {},
): LayoutParagraph {
	return { kind: 'paragraph', id, runs: [{ text }], ...extra };
}

function boxOf(result: ReturnType<typeof layoutDocument>, id: string): LayoutParagraphBox {
	const box = result.pages
		.flatMap((p) => p.columns.flatMap((c) => c.blocks))
		.find((b): b is LayoutParagraphBox => b.kind === 'paragraph' && b.blockId === id);
	if (!box) throw new Error(`missing box ${id}`);
	return box;
}

function modelOf(
	blocks: Paragraph[],
	paragraphStyles?: DocumentModel['paragraphStyles'],
): DocumentModel {
	return {
		blocks,
		page: {
			width: page.widthPx,
			height: page.heightPx,
			marginTop: page.marginTopPx,
			marginRight: page.marginRightPx,
			marginBottom: page.marginBottomPx,
			marginLeft: page.marginLeftPx,
		},
		...(paragraphStyles ? { paragraphStyles } : {}),
		warnings: [],
	};
}

describe('w:framePr parsing and style inheritance', () => {
	it('parses all non-drop-cap w:framePr attributes on direct paragraph properties', () => {
		const props = parseDirectParagraphProperties(
			pPrXml(
				'<w:framePr w:w="2880" w:h="1440" w:hRule="atLeast" w:x="360" w:y="-180" w:xAlign="right" w:yAlign="top" w:hAnchor="margin" w:vAnchor="page" w:hSpace="180" w:vSpace="90" w:wrap="notBeside"/>',
			),
		);
		expect(props.dropCap).toBeUndefined();
		expect(props.frame).toEqual({
			widthTwips: 2880,
			heightTwips: 1440,
			heightRule: 'atLeast',
			xTwips: 360,
			yTwips: -180,
			xAlign: 'right',
			yAlign: 'top',
			hAnchor: 'margin',
			vAnchor: 'page',
			hSpaceTwips: 180,
			vSpaceTwips: 90,
			wrap: 'notBeside',
		});
	});

	it('keeps drop-cap w:framePr on dropCap and leaves frame undefined', () => {
		const props = parseDirectParagraphProperties(
			pPrXml('<w:framePr w:dropCap="drop" w:lines="4" w:hSpace="144" w:vAnchor="text" w:x="80"/>'),
		);
		expect(props.dropCap).toEqual({ style: 'drop', lines: 4, distanceTwips: 144 });
		expect(props.frame).toBeUndefined();
	});

	it('ignores empty w:framePr or bare dropCap="none"', () => {
		expect(parseDirectParagraphProperties(pPrXml('<w:framePr/>')).frame).toBeUndefined();
		expect(
			parseDirectParagraphProperties(pPrXml('<w:framePr w:dropCap="none"/>')).frame,
		).toBeUndefined();
	});

	it('parses w:framePr in paragraph styles and merges inherited with direct overrides', () => {
		const catalog = parseParagraphStyleCatalog(`
			<w:styles xmlns:w="${W_NS}">
				<w:style w:type="paragraph" w:styleId="Sidebar">
					<w:pPr>
						<w:framePr w:w="3000" w:hAnchor="margin" w:vAnchor="margin" w:xAlign="right" w:hSpace="150" w:wrap="around"/>
					</w:pPr>
				</w:style>
				<w:style w:type="paragraph" w:styleId="SidebarTight">
					<w:basedOn w:val="Sidebar"/>
					<w:pPr>
						<w:framePr w:hSpace="60"/>
					</w:pPr>
				</w:style>
			</w:styles>
		`);
		const p: Paragraph = {
			type: 'paragraph',
			id: 'p1',
			style: 'SidebarTight',
			runs: [{ text: 'note' }],
			frame: { yTwips: signedTwips(750) },
		};
		const resolved = resolveParagraphFormatting(p, catalog);
		expect(resolved.frame).toEqual({
			widthTwips: 3000,
			hAnchor: 'margin',
			vAnchor: 'margin',
			xAlign: 'right',
			hSpaceTwips: 60,
			wrap: 'around',
			yTwips: 750,
		});
	});

	it('preserves w:framePr XML inside first() helper check', () => {
		const el = pPrXml('<w:framePr w:w="1500" w:wrap="none"/>');
		expect(first(el, 'framePr')).toBeDefined();
		expect(parseDirectParagraphProperties(el).frame).toEqual({ widthTwips: 1500, wrap: 'none' });
	});
});

describe('text frame positioning', () => {
	it('positions frames relative to page, margin, and text anchors with x/y offsets', () => {
		const makeFrame = (extra: Partial<LayoutTextFrame>): LayoutTextFrame => ({
			widthPx: 200,
			relativeFromH: 'column',
			offsetXPx: 0,
			relativeFromV: 'paragraph',
			offsetYPx: 0,
			hSpacePx: 0,
			vSpacePx: 0,
			wrap: 'none',
			...extra,
		});
		const blocks: LayoutBlock[] = [
			paragraph('lead', 'lead'),
			paragraph('onPage', 'page', {
				textFrame: makeFrame({
					relativeFromH: 'page',
					offsetXPx: 40,
					relativeFromV: 'page',
					offsetYPx: 50,
				}),
			}),
			paragraph('onMargin', 'margin', {
				textFrame: makeFrame({
					relativeFromH: 'margin',
					offsetXPx: 30,
					relativeFromV: 'margin',
					offsetYPx: 70,
				}),
			}),
			paragraph('onText', 'text', {
				textFrame: makeFrame({
					relativeFromH: 'column',
					offsetXPx: 15,
					relativeFromV: 'paragraph',
					offsetYPx: 10,
				}),
			}),
		];
		const result = layoutDocument({ sections: [{ page, blocks }] }, measurer);
		// Page coordinates: page left = 0, marginLeft = 100 -> box.xPx = 40 - 100 = -60; box.yPx = 50 - 100 = -50.
		const onPage = boxOf(result, 'onPage');
		expect(onPage.xPx).toBe(-60);
		expect(onPage.yPx).toBe(-50);
		expect(onPage.widthPx).toBe(200);
		// Margin coordinates: marginLeft = 100, marginTop = 100 -> box.xPx = 30, box.yPx = 70.
		const onMargin = boxOf(result, 'onMargin');
		expect(onMargin.xPx).toBe(30);
		expect(onMargin.yPx).toBe(70);
		// Text/paragraph coordinates: lead is 20px tall -> cursor.y = 20; +10 offset -> yPx = 30, xPx = 15.
		const onText = boxOf(result, 'onText');
		expect(onText.xPx).toBe(15);
		expect(onText.yPx).toBe(30);
	});

	it('honours xAlign and yAlign (left, center, right, top, bottom)', () => {
		const result = layoutDocument(
			{
				sections: [
					{
						page,
						blocks: [
							paragraph('centered', 'hi', {
								textFrame: {
									widthPx: 200,
									heightPx: 100,
									heightRule: 'exact',
									relativeFromH: 'margin',
									alignH: 'center',
									offsetXPx: 0,
									relativeFromV: 'margin',
									alignV: 'bottom',
									offsetYPx: 0,
									hSpacePx: 0,
									vSpacePx: 0,
									wrap: 'none',
								},
							}),
							paragraph('rightTop', 'hi', {
								textFrame: {
									widthPx: 150,
									relativeFromH: 'page',
									alignH: 'right',
									offsetXPx: 0,
									relativeFromV: 'page',
									alignV: 'top',
									offsetYPx: 0,
									hSpacePx: 0,
									vSpacePx: 0,
									wrap: 'none',
								},
							}),
						],
					},
				],
			},
			measurer,
		);
		// Margin area width = 800 - 100 - 100 = 600 -> centered 200px box starts at xPx = 200.
		// Margin area height = 1000 - 100 - 100 = 800 -> bottom-aligned 100px box starts at yPx = 700.
		const centered = boxOf(result, 'centered');
		expect(centered.xPx).toBe(200);
		expect(centered.yPx).toBe(700);
		// Page right-aligned 150px box: page x = 800 - 150 = 650 -> column-relative xPx = 550; page top y = 0 -> yPx = -100.
		const rightTop = boxOf(result, 'rightTop');
		expect(rightTop.xPx).toBe(550);
		expect(rightTop.yPx).toBe(-100);
	});

	it('auto-sizes width from content when w is omitted and honours hRule (auto, atLeast, exact)', () => {
		const result = layoutDocumentModel(
			modelOf([
				{
					type: 'paragraph',
					id: 'autoW',
					runs: [{ text: 'hello' }],
					frame: { hAnchor: 'margin', vAnchor: 'margin', wrap: 'none' },
				},
				{
					type: 'paragraph',
					id: 'atLeastH',
					runs: [{ text: 'short' }],
					frame: {
						widthTwips: twips(3000),
						heightTwips: twips(1500),
						heightRule: 'atLeast',
						hAnchor: 'margin',
						vAnchor: 'margin',
						wrap: 'around',
					},
				},
				{
					type: 'paragraph',
					id: 'exactH',
					runs: [{ text: 'line1' }],
					frame: {
						widthTwips: twips(3000),
						heightTwips: twips(600),
						heightRule: 'exact',
						hAnchor: 'margin',
						vAnchor: 'margin',
						wrap: 'around',
					},
				},
			]),
			measurer,
		);
		// 'hello' = 5 chars * 10px = 50px auto width.
		expect(boxOf(result, 'autoW').widthPx).toBe(50);
		// atLeast 1500 twips = 100px > 20px content height.
		expect(boxOf(result, 'atLeastH').floating?.frameHeightPx).toBe(100);
		// exact 600 twips = 40px.
		expect(boxOf(result, 'exactH').floating?.frameHeightPx).toBe(40);
	});
});

describe('consecutive framed paragraphs and wrapping', () => {
	it('groups consecutive paragraphs with matching framePr into one frame and wraps text around the whole frame', () => {
		const sharedFrame: ParagraphFrame = {
			widthTwips: twips(3000), // 200px
			hAnchor: 'margin',
			vAnchor: 'text',
			hSpaceTwips: twips(150), // 10px
			vSpaceTwips: twips(0),
			wrap: 'around',
		};
		const result = layoutDocumentModel(
			modelOf([
				{ type: 'paragraph', id: 'f1', runs: [{ text: 'first line' }], frame: sharedFrame },
				{ type: 'paragraph', id: 'f2', runs: [{ text: 'second line' }], frame: { ...sharedFrame } },
				{ type: 'paragraph', id: 'body', runs: [{ text: words }] },
			]),
			measurer,
		);
		const f1 = boxOf(result, 'f1');
		const f2 = boxOf(result, 'f2');
		const body = boxOf(result, 'body');
		// Grouped into one frame: f2 stacks directly below f1, sharing xPx and widthPx.
		expect(f1.xPx).toBe(0);
		expect(f1.yPx).toBe(0);
		expect(f2.xPx).toBe(0);
		expect(f2.yPx).toBe(f1.heightPx);
		expect(f1.floating?.frameHeightPx).toBe(f1.heightPx + f2.heightPx);
		expect(f2.floating?.frameHeightPx).toBeUndefined();
		// Body starts at y = 0 (frame took no flow room) and wraps beside both f1 and f2 (200 + 10 = 210px).
		expect(body.yPx).toBe(0);
		expect(body.lines[0]?.fragments[0]?.xPx).toBeGreaterThanOrEqual(210);
		expect(body.lines[1]?.fragments[0]?.xPx).toBeGreaterThanOrEqual(210);
		// Once past the 40px combined frame height, body text returns to the left margin (x = 0).
		expect(body.lines[2]?.fragments[0]?.xPx).toBe(0);
	});

	it('keeps consecutive paragraphs with different framePr in separate frames', () => {
		const result = layoutDocumentModel(
			modelOf([
				{
					type: 'paragraph',
					id: 'left',
					runs: [{ text: 'left' }],
					frame: { widthTwips: twips(1500), hAnchor: 'margin', xAlign: 'left', wrap: 'none' },
				},
				{
					type: 'paragraph',
					id: 'right',
					runs: [{ text: 'right' }],
					frame: { widthTwips: twips(1500), hAnchor: 'margin', xAlign: 'right', wrap: 'none' },
				},
			]),
			measurer,
		);
		const left = boxOf(result, 'left');
		const right = boxOf(result, 'right');
		expect(left.xPx).toBe(0);
		expect(right.xPx).toBe(500);
		expect(left.yPx).toBe(0);
		expect(right.yPx).toBe(0);
	});

	it('pushes body text below the frame when wrap="notBeside" and leaves it unwrapped when wrap="none"', () => {
		const notBeside = layoutDocumentModel(
			modelOf([
				{
					type: 'paragraph',
					id: 'frame',
					runs: [{ text: 'box' }],
					frame: {
						widthTwips: twips(3000), // 200px
						hAnchor: 'margin',
						vAnchor: 'text',
						vSpaceTwips: twips(150), // 10px
						wrap: 'notBeside',
					},
				},
				{ type: 'paragraph', id: 'body', runs: [{ text: words }] },
			]),
			measurer,
		);
		const bodyBelow = boxOf(notBeside, 'body');
		// Frame is 20px tall + 10px vSpace below -> first body line starts at y >= 30 and x = 0.
		expect(bodyBelow.yPx + bodyBelow.lines[0]!.yPx).toBeGreaterThanOrEqual(30);
		expect(bodyBelow.lines[0]!.fragments[0]!.xPx).toBe(0);

		const none = layoutDocumentModel(
			modelOf([
				{
					type: 'paragraph',
					id: 'frame',
					runs: [{ text: 'box' }],
					frame: {
						widthTwips: twips(3000),
						hAnchor: 'margin',
						vAnchor: 'text',
						wrap: 'none',
					},
				},
				{ type: 'paragraph', id: 'body', runs: [{ text: words }] },
			]),
			measurer,
		);
		const bodyThrough = boxOf(none, 'body');
		expect(bodyThrough.yPx).toBe(0);
		expect(bodyThrough.lines[0]!.yPx).toBe(0);
		expect(bodyThrough.lines[0]!.fragments[0]!.xPx).toBe(0);
	});

	it('does not regress drop-cap folding when text frame layout is enabled', () => {
		const result = layoutDocumentModel(
			modelOf([
				{
					type: 'paragraph',
					id: 'cap',
					runs: [{ text: 'O', fontSize: 55 }],
					dropCap: { style: 'drop', lines: 3 },
				},
				{ type: 'paragraph', id: 'body', runs: [{ text: 'nce upon a time' }] },
			]),
			measurer,
		);
		const blocks = result.pages[0]!.columns[0]!.blocks;
		expect(blocks).toHaveLength(1);
		const body = boxOf(result, 'body');
		expect(body.floating).toBeUndefined();
		expect(body.lines[0]!.fragments.map((f) => f.text).join('')).toBe('Once upon a time');
	});
});
