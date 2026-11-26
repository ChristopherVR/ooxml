import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from './xml';
import { updateDocPr } from './write-drawing-patch';
import {
	alignDrawingBox,
	alignDrawingToPage,
	createDocument,
	distributeDrawingBoxes,
	drawingMovement,
	drawingPosition,
	drawingStackOrder,
	groupDrawingShapes,
	loadDocx,
	newShape,
	placeDrawingBox,
	rotateDrawing,
	saveDocx,
	setDrawingWrap,
	ungroupDrawingShapes,
	type DrawingPosition,
	type InlineImage,
	type Paragraph,
	type PicturePlacement,
} from './index';

const rect = (x: number, y: number, width = 64, height = 43) => ({ x, y, width, height });
const shape = (): InlineImage => ({
	relId: '',
	partName: '',
	contentType: 'application/octet-stream',
	widthPx: 64,
	heightPx: 43,
	unsupported: 'Shape',
	shape: newShape('rect'),
});
const images = (block: Paragraph) => block.runs.flatMap((run) => (run.image ? [run.image] : []));
const xmlOf = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');

describe('Word drawing arrangement', () => {
	it('edits only the selected drawing description when a text box holds another drawing', () => {
		const ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
		const xml = parseXml(`<drawing xmlns:wp="${ns}"><wp:anchor><wp:docPr descr="Outside"/>
			<story><wp:inline><wp:docPr descr="Inside"/></wp:inline></story></wp:anchor></drawing>`);
		updateDocPr(xml.documentElement!, 'New description', undefined);
		const descriptions = Array.from(xml.getElementsByTagNameNS(ns, 'docPr'), (node) =>
			node.getAttribute('descr'),
		);
		expect(descriptions).toEqual(['New description', 'Inside']);
	});
	it('writes all nine positions as horizontal and vertical margin alignments', async () => {
		const model = createDocument();
		const block = model.blocks[0];
		if (block?.type !== 'paragraph') throw new Error('Missing paragraph');
		const positions: DrawingPosition[] = [
			'top-left',
			'top-center',
			'top-right',
			'center-left',
			'center-center',
			'center-right',
			'bottom-left',
			'bottom-center',
			'bottom-right',
		];
		block.runs = positions.map((position) => ({
			text: '',
			image: { ...shape(), anchored: true, placement: drawingPosition(shape(), position) },
		}));
		const bytes = await saveDocx(model);
		const xml = await xmlOf(bytes);
		const reopened = (await loadDocx(bytes)).model.blocks[0];
		if (reopened?.type !== 'paragraph') throw new Error('Missing paragraph');
		expect(images(reopened).map(({ placement }) => [placement?.alignV, placement?.align])).toEqual(
			positions.map((position) => position.split('-')),
		);
		expect(xml.match(/<wp:positionH relativeFrom="margin">/g)).toHaveLength(9);
		expect(xml.match(/<wp:positionV relativeFrom="margin">/g)).toHaveLength(9);
		expect(xml.match(/<wp:wrapSquare wrapText="bothSides"\/>/g)).toHaveLength(9);
	});

	it('patches a kept anchor, order, stacking layer, rotation and description as one round trip', async () => {
		const model = createDocument();
		const block = model.blocks[0];
		if (block?.type !== 'paragraph') throw new Error('Missing paragraph');
		block.runs = [{ text: '', image: setDrawingWrap(shape(), 'square') }];
		const opened = await loadDocx(await saveDocx(model));
		const paragraph = opened.model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Missing paragraph');
		const image = images(paragraph)[0]!;
		const back = setDrawingWrap(image, 'behind');
		const fixed = drawingMovement(back.placement!, 'fix', 140, 60);
		paragraph.runs[0]!.image = {
			...back,
			placement: { ...alignDrawingToPage(fixed, 'page', 'right'), relativeHeight: 12 },
			shape: { ...image.shape!, ...rotateDrawing({}, 'left'), flipH: true },
			altText: 'An accessible shape',
		};
		const bytes = await opened.save(opened.model);
		const xml = await xmlOf(bytes);
		expect(xml).toContain('relativeHeight="12" behindDoc="1"');
		expect(xml).toContain('<wp:positionH relativeFrom="page"><wp:align>right</wp:align>');
		expect(xml).toContain('<wp:positionV relativeFrom="page"><wp:posOffset>1333500</wp:posOffset>');
		expect(xml).toContain('<wp:wrapNone/>');
		expect(xml).toContain('rot="16200000" flipH="1"');
		const reopened = (await loadDocx(bytes)).model.blocks[0];
		if (reopened?.type !== 'paragraph') throw new Error('Missing paragraph');
		const edited = images(reopened)[0]!;
		expect(edited).toMatchObject({
			placement: {
				wrap: 'none',
				behindText: true,
				relativeHeight: 12,
				relativeFromV: 'page',
				offsetYPx: 140,
			},
			shape: { rotation: 270, flipH: true },
			altText: 'An accessible shape',
		});
		expect(drawingMovement(edited.placement!, 'move', 140, 60)).toMatchObject({
			relativeFromV: 'paragraph',
			offsetYPx: 80,
		});
	});

	it('writes a native group and saves its members as individual shapes after ungrouping', async () => {
		const model = createDocument();
		const block = model.blocks[0];
		if (block?.type !== 'paragraph') throw new Error('Missing paragraph');
		const members = [
			{ box: rect(110, 140), shape: newShape('rect') },
			{ box: rect(260, 175, 95, 60), shape: newShape('ellipse') },
		];
		const group = groupDrawingShapes(members)!;
		const image = {
			...shape(),
			widthPx: group.box.width,
			heightPx: group.box.height,
			shape: group.shape,
			anchored: true,
			placement: placeDrawingBox({ wrap: 'square' }, group.box),
		};
		block.runs = [{ text: '', image }];
		const bytes = await saveDocx(model);
		const xml = await xmlOf(bytes);
		expect(xml).toContain('<wpg:wgp');
		expect(xml).toContain('<wpg:cNvGrpSpPr');
		expect(xml).toContain('<a:chExt');
		const opened = await loadDocx(bytes);
		const paragraph = opened.model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Missing paragraph');
		const groupImage = images(paragraph)[0]!;
		const separated = ungroupDrawingShapes(group.box, groupImage.shape!);
		expect(separated.map(({ box }) => box)).toEqual(members.map(({ box }) => box));
		paragraph.runs = separated.map(({ box, shape: member }) => ({
			text: '',
			image: {
				...groupImage,
				shape: member,
				widthPx: box.width,
				heightPx: box.height,
				placement: placeDrawingBox(groupImage.placement!, box),
			},
		}));
		const ungrouped = await opened.save(opened.model);
		const ungroupedXml = await xmlOf(ungrouped);
		expect(ungroupedXml).not.toContain('<wpg:wgp');
		expect(ungroupedXml.match(/<wp:anchor /g)).toHaveLength(2);
		const reopened = (await loadDocx(ungrouped)).model.blocks[0];
		if (reopened?.type !== 'paragraph') throw new Error('Missing paragraph');
		expect(images(reopened).map((drawing) => drawing.shape?.preset)).toEqual(['rect', 'ellipse']);
	});

	it('places only the requested axes in the existing drawing frames', () => {
		const placement: PicturePlacement = {
			wrap: 'none',
			relativeFrom: 'column',
			align: 'center',
			relativeFromV: 'paragraph',
			alignV: 'bottom',
		};
		expect(placeDrawingBox(placement, rect(140, 205), { x: 96 })).toEqual({
			wrap: 'none',
			relativeFrom: 'column',
			offsetXPx: 44,
			relativeFromV: 'paragraph',
			alignV: 'bottom',
		});
		expect(placeDrawingBox(placement, rect(140, 205), { y: 220 })).toEqual({
			wrap: 'none',
			relativeFrom: 'column',
			align: 'center',
			relativeFromV: 'paragraph',
			offsetYPx: -15,
		});
		expect(placeDrawingBox(placement, rect(140, 205), { x: 96, y: 180 })).toEqual({
			wrap: 'none',
			relativeFrom: 'column',
			offsetXPx: 44,
			relativeFromV: 'paragraph',
			offsetYPx: 25,
		});
	});

	it('aligns and distributes unequal boxes and keeps a stable stack', () => {
		const boxes = [rect(100, 45, 12), rect(0, 30, 30), rect(60, 15, 20)];
		expect(distributeDrawingBoxes(boxes, rect(0, 0, 112), 'x').map(({ x }) => x)).toEqual([
			100, 0, 55,
		]);
		expect(alignDrawingBox(boxes[0]!, rect(0, 5, 112, 80), 'middle').y).toBe(23.5);
		expect(drawingStackOrder([30, 10, 20], 1, 'forward')).toEqual([
			251658242, 251658241, 251658240,
		]);
		expect(drawingStackOrder([30, 10, 20], 0, 'forward')).toBeUndefined();
	});

	it('writes table cell text direction and edits it without losing cell properties', async () => {
		const model = createDocument();
		const paragraph = model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Missing paragraph');
		model.blocks = [
			{
				type: 'table',
				id: 'table1',
				rows: [[{ paragraphs: [paragraph], verticalAlign: 'bottom', textDirection: 'tbRl' }]],
			},
		];
		const opened = await loadDocx(await saveDocx(model));
		const table = opened.model.blocks[0];
		if (table?.type !== 'table') throw new Error('Missing table');
		const cell = table.rows[0]![0]!;
		expect(cell.textDirection).toBe('tbRl');
		cell.textDirection = 'btLr';
		const saved = await opened.save(opened.model);
		expect(await xmlOf(saved)).toMatch(
			/<w:textDirection w:val="btLr"\/><w:vAlign w:val="bottom"\/>/,
		);
	});

	it('changes the direction of a kept text box without rebuilding its text', async () => {
		const model = createDocument();
		const block = model.blocks[0];
		if (block?.type !== 'paragraph') throw new Error('Missing paragraph');
		block.runs = [
			{
				text: '',
				image: {
					relId: '',
					partName: '',
					contentType: 'application/octet-stream',
					widthPx: 144,
					heightPx: 64,
					unsupported: 'Text box',
					textBoxEditable: true,
					textBoxText: ['Keep this text'],
				},
			},
		];
		const opened = await loadDocx(await saveDocx(model));
		const paragraph = opened.model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Missing paragraph');
		const image = images(paragraph)[0]!;
		image.textBoxContent = {
			...image.textBoxContent,
			blocks: image.textBoxContent?.blocks ?? [],
			vertical: 'vert270',
		};
		const bytes = await opened.save(opened.model);
		const xml = await xmlOf(bytes);
		expect(xml).toMatch(/<wps:bodyPr[^>]*vert="vert270"/);
		expect(xml).toContain('<w:t>Keep this text</w:t>');
		const reopened = (await loadDocx(bytes)).model.blocks[0];
		if (reopened?.type !== 'paragraph') throw new Error('Missing paragraph');
		expect(images(reopened)[0]?.textBoxContent?.vertical).toBe('vert270');
	});
});
