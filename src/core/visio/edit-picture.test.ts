import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { VISIO_PICTURE_MAX_BYTES, type VisioPictureInsertEdit } from './edit-metadata-commands';
import { cell, fixture, shape } from './test-fixtures';

const png = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
		'base64',
	),
);
const gif = Uint8Array.from(
	Buffer.from('R0lGODlhAQABAIAAAP///wAAACwAAAAAAQABAAACAkQBADs=', 'base64'),
);
const insert = (patch: Partial<VisioPictureInsertEdit> = {}): VisioPictureInsertEdit => ({
	type: 'insert-picture',
	pageId: '0',
	shapeId: '1',
	x: 4.25,
	y: 5.5,
	width: 2,
	height: 1,
	image: png,
	...patch,
});
const text = async (bytes: Uint8Array, path: string) =>
	new TextDecoder().decode(await (await VisioPackage.open(bytes)).readBytes(path));

describe('picture insertion', () => {
	it('embeds a PNG as a Foreign shape with a media part, page relationship and content type', async () => {
		const result = await editVsdx(await createVsdx(), [insert()]);
		expect(result.changedParts).toContain('visio/media/image1.png');
		const types = await text(result.bytes, '[Content_Types].xml');
		expect(types).toContain('<Default Extension="png" ContentType="image/png"/>');
		const rels = await text(result.bytes, 'visio/pages/_rels/page1.xml.rels');
		expect(rels).toMatch(
			/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image" Target="\.\.\/media\/image1\.png"/,
		);
		const page = await text(result.bytes, 'visio/pages/page1.xml');
		expect(page).toContain('Type="Foreign"');
		expect(page).toContain('<ForeignData ForeignType="Bitmap" CompressionType="PNG">');
		expect(page).toContain('N="ImgWidth" V="2" F="Width*1"');
		const shape = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		expect(shape).toMatchObject({ id: '1', kind: 'foreign', width: 2, height: 1 });
		expect(shape.image).toMatchObject({
			mimeType: 'image/png',
			pixelWidth: 1,
			width: 2,
			height: 1,
		});
		expect(shape.geometry.every((path) => !path.stroke && !path.fill)).toBe(true);
	});
	it('moves, resizes and deletes the picture with the existing edits', async () => {
		const inserted = (await editVsdx(await createVsdx(), [insert()])).bytes;
		const moved = await editVsdx(inserted, [
			{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 3 },
		]);
		const resized = await editVsdx(moved.bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 2 },
		]);
		const image = (await parseVsdx(resized.bytes)).pages[0]!.shapes[0]!;
		expect(image.rotation).toMatchObject({ pinX: 2, pinY: 3 });
		expect(image.image).toMatchObject({ width: 4, height: 2 });
		const removed = await editVsdx(resized.bytes, [
			{ type: 'delete-shape', pageId: '0', shapeId: '1' },
		]);
		expect((await parseVsdx(removed.bytes)).pages[0]!.shapes).toHaveLength(0);
	});
	it('chooses free part names, relationship IDs and adds a second format default', async () => {
		const first = (await editVsdx(await createVsdx(), [insert()])).bytes;
		const second = await editVsdx(first, [insert({ shapeId: '2', image: gif })]);
		expect(second.changedParts).toContain('visio/media/image2.gif');
		const rels = await text(second.bytes, 'visio/pages/_rels/page1.xml.rels');
		expect(rels).toContain('Id="rId1"');
		expect(rels).toContain('Id="rId2"');
		expect(await text(second.bytes, '[Content_Types].xml')).toContain('Extension="gif"');
		const shapes = (await parseVsdx(second.bytes)).pages[0]!.shapes;
		expect(shapes.map((item) => item.image?.mimeType)).toEqual(['image/png', 'image/gif']);
	});
	it('adds to an existing page relationship part without disturbing other relationships', async () => {
		const withMedia = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${shape('1')}</Shapes>` }],
			edit: (zip) => {
				zip.file('visio/media/image1.png', png);
				zip.file(
					'visio/pages/_rels/page1.xml.rels',
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
				);
			},
		});
		const result = await editVsdx(withMedia, [insert({ shapeId: '2' })]);
		const rels = await text(result.bytes, 'visio/pages/_rels/page1.xml.rels');
		expect(rels).toContain('Id="rId1"');
		expect(rels).toContain('Id="rId2"');
		expect(rels).toContain('Target="../media/image2.png"');
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[1]!.image?.mimeType).toBe('image/png');
	});
	it('rejects invalid, unsupported and oversized images, ID collisions and mixed transactions', async () => {
		const blank = await createVsdx();
		const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
		await expect(editVsdx(blank, [insert({ image: svg })])).rejects.toMatchObject({
			code: 'INVALID_PICTURE',
		});
		await expect(editVsdx(blank, [insert({ image: png.slice(0, 20) })])).rejects.toMatchObject({
			code: 'INVALID_PICTURE',
		});
		await expect(
			editVsdx(blank, [insert({ image: new Uint8Array(VISIO_PICTURE_MAX_BYTES + 1) })]),
		).rejects.toMatchObject({ code: 'LIMIT_PICTURE' });
		await expect(editVsdx(blank, [insert({ width: 0 })])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
		const occupied = await fixture();
		await expect(editVsdx(occupied, [insert()])).rejects.toMatchObject({
			code: 'INVALID_SHAPE_ID',
		});
		const referenced = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${shape('1', cell('Width', 1, 'Sheet.2!Width'))}</Shapes>` },
			],
		});
		await expect(editVsdx(referenced, [insert({ shapeId: '2' })])).rejects.toMatchObject({
			code: 'INVALID_SHAPE_ID',
		});
		await expect(
			editVsdx(blank, [insert(), { type: 'move-shape', pageId: '0', shapeId: '1', x: 1, y: 1 }]),
		).rejects.toMatchObject({ code: 'EDIT_MIXED_PICTURE_TRANSACTION' });
		await expect(editVsdx(blank, [insert({ pageId: '9' })])).rejects.toMatchObject({
			code: 'EDIT_TARGET_NOT_FOUND',
		});
	});
	it('does not share caller-owned image bytes with the transaction', async () => {
		const bytes = new Uint8Array(png);
		const pending = editVsdx(await createVsdx(), [insert({ image: bytes })]);
		bytes.fill(0);
		const result = await pending;
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.image?.mimeType).toBe('image/png');
	});
});
