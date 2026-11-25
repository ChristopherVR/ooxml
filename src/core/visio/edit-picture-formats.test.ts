import { convertMetafileToSvgTree } from 'emf-converter';
import { describe, expect, it } from 'vitest';
import { decodePng } from '../drawingml/raster-decode';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import type { VisioPictureInsertEdit } from './edit-metadata-commands';
import { emf, record, setWord } from './emf-admission.test-fixtures';
import { VisioPackage } from './package';
import { parseVsdx } from './parser';
import { bmpPicture, svgPicture, tiffPicture } from './picture-test-fixtures';

const insertPicture = (image: Uint8Array): VisioPictureInsertEdit => ({
	type: 'insert-picture',
	pageId: '0',
	shapeId: '1',
	x: 1,
	y: 2,
	width: 2,
	height: 1,
	image,
});
const brushEmf = () =>
	emf([record(39, [1, 0, 255, 0]), record(37, [1]), record(43, [10, 10, 90, 90])]);
const decode = new TextDecoder();

describe('Visio picture format round trips', () => {
	it.each([
		['bmp', 'image/bmp', '', bmpPicture()],
		['tiff', 'image/tiff', 'TIFF', tiffPicture(1, false)],
		['tiff', 'image/tiff', 'TIFF', tiffPicture(5)],
		['tiff', 'image/tiff', 'TIFF', tiffPicture(32773)],
		['svg', 'image/svg+xml', '', svgPicture()],
	] as const)(
		'writes and reopens %s media as an editable Foreign bitmap frame',
		async (extension, mime, compression, source) => {
			const inserted = await editVsdx(await createVsdx(), [insertPicture(source)]);
			const part = 'visio/media/image1.' + extension;
			expect(inserted.changedParts).toContain(part);
			const pkg = await VisioPackage.open(inserted.bytes);
			expect(decode.decode(await pkg.readBytes('[Content_Types].xml'))).toContain(
				'Extension="' + extension + '" ContentType="' + mime + '"',
			);
			expect(decode.decode(await pkg.readBytes('visio/pages/page1.xml'))).toContain(
				compression
					? 'ForeignType="Bitmap" CompressionType="' + compression + '"'
					: '<ForeignData ForeignType="Bitmap">',
			);
			const resized = await editVsdx(inserted.bytes, [
				{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 2 },
			]);
			const image = (await parseVsdx(resized.bytes)).pages[0]?.shapes[0]?.image;
			expect(image).toMatchObject({
				mimeType: extension === 'tiff' ? 'image/png' : mime,
				pixelWidth: 2,
				pixelHeight: 1,
				width: 4,
				height: 2,
			});
			if (extension === 'tiff') {
				expect(decodePng(image!.bytes)?.rgb).toEqual(Uint8Array.of(255, 0, 0, 0, 255, 0));
			}
			const reopened = await VisioPackage.open(resized.bytes);
			expect(await reopened.readBytes(part)).toEqual(source);
		},
	);

	it('reopens an EMF using a brush record outside the previous conversion subset', async () => {
		const source = brushEmf();
		const inserted = await editVsdx(await createVsdx(), [insertPicture(source)]);
		const resized = await editVsdx(inserted.bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 2 },
		]);
		const pkg = await VisioPackage.open(resized.bytes);
		expect(await pkg.readBytes('visio/media/image1.emf')).toEqual(source);
		expect(decode.decode(await pkg.readBytes('visio/pages/page1.xml'))).toContain(
			'ForeignType="EnhMetaFile"',
		);
		const model = await parseVsdx(resized.bytes, { metafileConverter: convertMetafileToSvgTree });
		const foreign = model.pages[0]?.shapes[0]?.foreignVector;
		expect(foreign).toMatchObject({
			width: 4,
			height: 2,
			vector: { width: 100, height: 100 },
		});
		expect(foreign?.vector.items).toContainEqual(
			expect.objectContaining({
				kind: 'path',
				paint: expect.objectContaining({ fill: '#ff0000' }),
			}),
		);
	});

	it.each([
		['SVG script', svgPicture('<script/>')],
		['SVG event handler', svgPicture('<rect onload="void(0)"/>')],
		[
			'SVG external paint',
			svgPicture('<rect fill="url(https://example.test/file.svg#gradient)"/>'),
		],
		['truncated BMP', bmpPicture().slice(0, -1)],
		['truncated TIFF', tiffPicture().slice(0, -1)],
		['truncated EMF', brushEmf().slice(0, -1)],
		['invalid EMF record size', setWord(brushEmf(), 112, 0)],
	])('refuses %s before editing the package', async (_name, image) => {
		await expect(editVsdx(await createVsdx(), [insertPicture(image)])).rejects.toMatchObject({
			code: 'INVALID_PICTURE',
		});
	});
});
