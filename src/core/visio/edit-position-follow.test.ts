import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

describe('native position-following creation', () => {
	it('round trips pin formulas and refuses deletion of a referenced owner', async () => {
		const initial = await editVsdx(await createVsdx(), [
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 2, width: 1, height: 1 },
			{
				type: 'create-text-box',
				pageId: '0',
				shapeId: '2',
				x: 3,
				y: 1.2,
				width: 1,
				height: 0.2,
				text: 'Metric',
				followShapeId: '1',
			},
		]);
		const moved = await editVsdx(initial.bytes, [
			{ type: 'move-shape', pageId: '0', shapeId: '1', x: 4, y: 5 },
		]);
		const follower = (await parseVsdx(moved.bytes)).pages[0]!.shapes[1]!;
		expect(follower.rotation!.pinX).toBe(5);
		expect(follower.rotation!.pinY).toBeCloseTo(4.2);
		await expect(
			editVsdx(moved.bytes, [{ type: 'delete-shape', pageId: '0', shapeId: '1' }]),
		).rejects.toThrow(/referenc/i);
		await expect(
			editVsdx(
				moved.bytes,
				['1', '2'].map((shapeId) => ({ type: 'delete-shape' as const, pageId: '0', shapeId })),
			),
		).resolves.toBeDefined();
	});
	it.each(['2', 'invalid', '1)!Width', '4294967296'])(
		'rejects self references or invalid owner syntax %s',
		async (followShapeId) => {
			await expect(
				editVsdx(await createVsdx(), [
					{
						type: 'create-rectangle',
						pageId: '0',
						shapeId: '2',
						x: 1,
						y: 1,
						width: 1,
						height: 1,
						followShapeId,
					},
				]),
			).rejects.toThrow(/owner/);
		},
	);
	it('refuses missing and one-dimensional owners', async () => {
		const bytes = (
			await editVsdx(await createVsdx(), [
				{ type: 'create-line', pageId: '0', shapeId: '1', beginX: 1, beginY: 1, endX: 3, endY: 1 },
			])
		).bytes;
		for (const followShapeId of ['1', '3'])
			await expect(
				editVsdx(bytes, [
					{
						type: 'create-ellipse',
						pageId: '0',
						shapeId: '2',
						x: 1,
						y: 1,
						width: 1,
						height: 1,
						followShapeId,
					},
				]),
			).rejects.toThrow();
	});
});
