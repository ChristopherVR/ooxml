import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { parseVsdx } from './parser';
import { editVsdx, type VisioEdit } from './edit';

const png = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
		'base64',
	),
);
const picture = async (): Promise<Uint8Array> =>
	(
		await editVsdx(await createVsdx(), [
			{
				type: 'insert-picture',
				pageId: '0',
				shapeId: '1',
				x: 4.25,
				y: 5.5,
				width: 2,
				height: 1,
				image: png,
			} as VisioEdit,
		])
	).bytes;
const shapeOf = async (bytes: Uint8Array) =>
	(await parseVsdx(bytes)).pages[0]!.shapes.find((shape) => shape.id === '1')!;
const adjust = (patch: { brightness?: number; contrast?: number }): VisioEdit =>
	({ type: 'set-picture-adjust', pageId: '0', shapeId: '1', ...patch }) as VisioEdit;
const crop = (mode: 'square' | 'reset'): VisioEdit =>
	({ type: 'crop-picture', pageId: '0', shapeId: '1', crop: mode }) as VisioEdit;

describe('picture brightness and contrast', () => {
	it('writes the Brightness and Contrast cells and reports them when they are not neutral', async () => {
		const bytes = (await editVsdx(await picture(), [adjust({ brightness: 0.7, contrast: 0.3 })]))
			.bytes;
		expect((await shapeOf(bytes)).image).toMatchObject({ brightness: 0.7, contrast: 0.3 });
		const neutral = (await editVsdx(bytes, [adjust({ brightness: 0.5, contrast: 0.5 })])).bytes;
		const image = (await shapeOf(neutral)).image!;
		expect(image.brightness).toBeUndefined();
		expect(image.contrast).toBeUndefined();
	});

	it('rejects values outside 0 to 1 and an empty adjustment', async () => {
		await expect(editVsdx(await picture(), [adjust({ brightness: 1.5 })])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
		await expect(editVsdx(await picture(), [adjust({})])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
	});
});

describe('picture crop', () => {
	it('crops the longer side equally so the picture is square, keeping its scale and place', async () => {
		const shape = await shapeOf(
			await editVsdx(await picture(), [crop('square')]).then((r) => r.bytes),
		);
		expect(shape.width).toBeCloseTo(1);
		expect(shape.height).toBeCloseTo(1);
		expect(shape.image?.x).toBeCloseTo(-0.5);
		expect(shape.image).toMatchObject({ width: 2, height: 1 });
	});

	it('restores the whole picture on reset', async () => {
		const squared = (await editVsdx(await picture(), [crop('square')])).bytes;
		const reset = await shapeOf((await editVsdx(squared, [crop('reset')])).bytes);
		expect(reset.width).toBeCloseTo(2);
		expect(reset.height).toBeCloseTo(1);
		expect(reset.image).toMatchObject({ x: 0, y: 0, width: 2, height: 1 });
	});

	it('refuses a shape that is not a picture', async () => {
		const bytes = (
			await editVsdx(await createVsdx(), [
				{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 1, width: 1, height: 1 },
			])
		).bytes;
		await expect(editVsdx(bytes, [crop('square')])).rejects.toMatchObject({
			code: 'EDIT_TARGET_NOT_FOUND',
		});
	});
});
