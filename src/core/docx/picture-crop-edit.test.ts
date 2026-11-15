import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	isValidPictureCrop,
	loadDocx,
	pictureCropOf,
	setPictureCrop,
	uncroppedPictureSize,
	type InlineImage,
	type Paragraph,
} from './index';
import { at } from './test-support/access';

const NS = [
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(' ');
// As Word 16 saved a 96 pt square picture after PictureFormat.CropLeft = 10 and CropTop = 5 (points).
const blipFill = (rect: string) =>
	`<pic:blipFill><a:blip r:embed="rId1"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip>${rect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>`;
const drawing = (cx: number, cy: number, rect: string) =>
	`<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>${blipFill(rect)}<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
const WORD_CROPPED = drawing(1092200, 1155700, '<a:srcRect l="10417" t="5208"/>');
const WORD_UNCROPPED = drawing(1219200, 1219200, '');

async function load(xml: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body><w:p><w:r>${xml}</w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', new Uint8Array([137, 80, 78, 71]));
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const picture = (loaded: Awaited<ReturnType<typeof loadDocx>>) =>
	at((at(loaded.model.blocks, 0) as Paragraph).runs, 0);
const none = { cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0 };

async function recrop(xml: string, crop: typeof none) {
	const loaded = await load(xml);
	const run = picture(loaded);
	run.image = setPictureCrop(run.image!, crop);
	const saved = await loaded.save(loaded.model);
	const out = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
	return { xml: out, image: picture(await loadDocx(saved)).image! };
}

describe('picture crop editing', () => {
	it('reads the crop and the whole size of a picture Word cropped', async () => {
		const image = picture(await load(WORD_CROPPED)).image!;
		expect(pictureCropOf(image)).toEqual({ ...none, cropLeft: 0.10417, cropTop: 0.05208 });
		const whole = uncroppedPictureSize(image);
		expect(whole.widthPx).toBeCloseTo(128, 0);
		expect(whole.heightPx).toBeCloseTo(128, 0);
	});

	it('crops like Word: the source rectangle changes and the shown size keeps its scale', async () => {
		const { xml, image } = await recrop(WORD_UNCROPPED, {
			...none,
			cropLeft: 0.10417,
			cropTop: 0.05208,
		});
		expect(xml).toContain(blipFill('<a:srcRect l="10417" t="5208"/>'));
		// Word wrote cx="1092200" cy="1155700" for the same crop (its extents are whole points).
		const [, cx, cy] = xml.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/)!;
		expect(Math.abs(Number(cx) - 1092200)).toBeLessThan(100);
		expect(Math.abs(Number(cy) - 1155700)).toBeLessThan(100);
		expect(xml).toContain(`<a:ext cx="${cx}" cy="${cy}"/>`);
		expect(pictureCropOf(image)).toMatchObject({ cropLeft: 0.10417, cropTop: 0.05208 });
	});

	it('changes and clears an existing crop, keeping the rest of the blip fill', async () => {
		const changed = await recrop(WORD_CROPPED, { ...none, cropRight: 0.25, cropBottom: 0.5 });
		expect(changed.xml).toContain(blipFill('<a:srcRect r="25000" b="50000"/>'));
		expect(changed.image.widthPx).toBeCloseTo(96, 0);
		expect(changed.image.heightPx).toBeCloseTo(64, 0);
		const cleared = await recrop(WORD_CROPPED, none);
		expect(cleared.xml).not.toMatch(/srcRect [ltrb]=/);
		expect(pictureCropOf(cleared.image)).toEqual(none);
		expect(cleared.image.widthPx).toBeCloseTo(128, 0);
	});

	it('refuses a crop that leaves nothing showing, and a drawing that is not a picture', () => {
		const image: InlineImage = {
			relId: 'r',
			partName: 'p',
			contentType: 'image/png',
			widthPx: 100,
			heightPx: 100,
			pictureFill: { kind: 'picture' } as NonNullable<InlineImage['pictureFill']>,
		};
		expect(isValidPictureCrop({ ...none, cropLeft: 0.6, cropRight: 0.4 })).toBe(false);
		expect(() => setPictureCrop(image, { ...none, cropTop: 0.5, cropBottom: 0.5 })).toThrow(
			RangeError,
		);
		const { pictureFill: _fill, ...shape } = image;
		expect(() => setPictureCrop(shape, none)).toThrow();
		expect(setPictureCrop(image, { ...none, cropLeft: -0.1 }).widthPx).toBeCloseTo(110);
	});
});
