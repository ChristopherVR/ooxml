import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	DRAWING_WRAPS,
	drawingWrapOf,
	loadDocx,
	setDrawingWrap,
	type InlineImage,
	type Paragraph,
} from './index';
import { at } from './test-support/access';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS = [
	`xmlns:w="${W}"`,
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(' ');
// The graphic and the inline and floating drawings below are as Word 16 saved them.
const GRAPHIC =
	'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="353617014" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1219200" cy="1219200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>';
const TAIL = `<wp:docPr id="353617014" name="Picture 1"/><wp:cNvGraphicFramePr/>${GRAPHIC}`;
const EXTENT = '<wp:extent cx="1219200" cy="1219200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>';
const INLINE = `<wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="25012050" wp14:editId="6025EB05">${EXTENT}${TAIL}</wp:inline>`;
const POLYGON =
	'<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21007"/><wp:lineTo x="18460" y="21007"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
const anchor = (wrap: string, behindDoc = '0') =>
	`<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="${behindDoc}" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="25012050" wp14:editId="137FC118"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>444500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>${EXTENT}${wrap}${TAIL}</wp:anchor>`;

async function load(drawing: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body><w:p><w:r><w:t>Before </w:t></w:r><w:r><w:drawing>${drawing}</w:drawing></w:r><w:r><w:t> after.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', new Uint8Array([137, 80, 78, 71]));
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const runs = (loaded: Awaited<ReturnType<typeof loadDocx>>) =>
	(at(loaded.model.blocks, 0) as Paragraph).runs;
const pictureOf = (loaded: Awaited<ReturnType<typeof loadDocx>>): InlineImage =>
	at(runs(loaded), 1).image!;

/** Applies `wrap`, saves, and returns the saved drawing XML and the reopened picture. */
async function rewrap(drawing: string, wrap: (typeof DRAWING_WRAPS)[number]) {
	const loaded = await load(drawing);
	at(runs(loaded), 1).image = setDrawingWrap(pictureOf(loaded), wrap);
	const saved = await loaded.save(loaded.model);
	const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
	const reopened = await loadDocx(saved);
	return { xml: xml.match(/<w:drawing>.*<\/w:drawing>/)![0], image: pictureOf(reopened) };
}

describe('Wrap Text on a picture', () => {
	it('maps each choice onto the model and back', () => {
		const image: InlineImage = {
			relId: 'r',
			partName: 'p',
			contentType: 'c',
			widthPx: 1,
			heightPx: 1,
		};
		for (const wrap of DRAWING_WRAPS) {
			const next = setDrawingWrap(image, wrap);
			expect(drawingWrapOf(next.anchored, next.placement)).toBe(wrap);
		}
		expect(setDrawingWrap(setDrawingWrap(image, 'square'), 'inline')).toEqual(image);
	});

	it('turns an inline picture into a floating one the way Word writes it', async () => {
		const { xml, image } = await rewrap(INLINE, 'square');
		const relativeHeight = 251658240 + (353617014 % 4096) * 1024;
		expect(xml).toBe(
			`<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${relativeHeight}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="25012050" wp14:editId="6025EB05"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
				`${EXTENT}<wp:wrapSquare wrapText="bothSides"/>${TAIL}</wp:anchor></w:drawing>`,
		);
		expect(image).toMatchObject({ anchored: true, placement: { wrap: 'square' } });
	});

	it('writes every choice from In Line with Text and reads it back', async () => {
		for (const wrap of DRAWING_WRAPS) {
			const { xml, image } = await rewrap(INLINE, wrap);
			expect(drawingWrapOf(image.anchored, image.placement)).toBe(wrap);
			expect(xml).toContain('wp14:anchorId="25012050"');
			expect(xml).toContain('<wp:docPr id="353617014" name="Picture 1"/><wp:cNvGraphicFramePr/>');
		}
		expect((await rewrap(INLINE, 'tight')).xml).toContain(
			'<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>',
		);
		expect((await rewrap(INLINE, 'behind')).xml).toMatch(/behindDoc="1".*<wp:wrapNone\/>/);
	});

	it('changes a floating wrap in place, keeping its position, distances and polygon', async () => {
		const tight = await rewrap(
			anchor(`<wp:wrapTight wrapText="left">${POLYGON}</wp:wrapTight>`),
			'through',
		);
		expect(tight.xml).toContain(`<wp:wrapThrough wrapText="left">${POLYGON}</wp:wrapThrough>`);
		expect(tight.xml).toContain('<wp:posOffset>444500</wp:posOffset>');
		expect(tight.xml).toContain('distL="114300"');
		expect(tight.xml).toContain('wp14:editId="137FC118"');
		const behind = await rewrap(anchor('<wp:wrapSquare wrapText="bothSides"/>'), 'behind');
		expect(behind.xml).toContain('behindDoc="1"');
		expect(behind.xml).toContain(`${EXTENT}<wp:wrapNone/><wp:docPr`);
		const front = await rewrap(anchor('<wp:wrapNone/>', '1'), 'front');
		expect(front.xml).toContain('behindDoc="0"');
		expect(front.image.placement?.wrap).toBe('none');
		expect(front.image.placement?.offsetXPx).toBeCloseTo(444500 / 9525, 0);
		const top = await rewrap(anchor('<wp:wrapNone/>'), 'topAndBottom');
		expect(top.xml).toContain(`${EXTENT}<wp:wrapTopAndBottom/><wp:docPr`);
	});

	it('turns a floating picture back into an inline one the way Word writes it', async () => {
		const { xml, image } = await rewrap(anchor('<wp:wrapSquare wrapText="bothSides"/>'), 'inline');
		expect(xml).toBe(
			`<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="25012050" wp14:editId="137FC118">${EXTENT}${TAIL}</wp:inline></w:drawing>`,
		);
		expect(image.anchored).toBeUndefined();
		expect(image.placement).toBeUndefined();
	});

	it('leaves the drawing untouched when the choice does not change', async () => {
		const source = anchor('<wp:wrapSquare wrapText="bothSides"/>');
		const { xml } = await rewrap(source, 'square');
		expect(xml).toBe(`<w:drawing>${source}</w:drawing>`);
	});

	it('writes a newly inserted picture that floats as an anchor', async () => {
		const loaded = await load(INLINE);
		runs(loaded).push({
			text: '',
			image: setDrawingWrap(
				{
					relId: '',
					partName: 'word/media/new.png',
					contentType: 'image/png',
					widthPx: 32,
					heightPx: 32,
				},
				'topAndBottom',
				{ offsetXPx: 10, offsetYPx: 5 },
			),
		});
		const media = new Map([
			[
				'word/media/new.png',
				{ bytes: new Uint8Array([137, 80, 78, 71]), contentType: 'image/png' },
			],
		]);
		const reopened = await loadDocx(await loaded.save(loaded.model, media));
		expect(runs(reopened).at(-1)?.image).toMatchObject({
			anchored: true,
			placement: { wrap: 'topAndBottom', offsetXPx: 10, offsetYPx: 5 },
		});
	});
});
