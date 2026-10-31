import { describe, expect, it } from 'vitest';
import { createVsdx } from '../create-document';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import type { VisioEdit } from '../edit-commands';
import { visioDataGraphicItems, visioSetDataGraphicItemEdits } from './data-graphic-items';
import {
	visioAutoSpaceDataGraphicEdits,
	visioDataGraphicConfigs,
	visioDataGraphicEdits,
	visioDataGraphicOwners,
	visioDataGraphicParts,
	visioPositionDataGraphicEdits,
	visioRemoveDataGraphicEdits,
} from './data-graphics';

const data = (name: string, value: string) => ({
	label: name,
	prompt: '',
	type: 'number' as const,
	format: '',
	value,
});
async function start(stencil = false): Promise<Uint8Array> {
	return (
		await editVsdx(await createVsdx(), [
			stencil
				? {
						type: 'drop-stencil-master' as const,
						pageId: '0',
						shapeId: '1',
						master: 'rectangle',
						x: 3,
						y: 4,
					}
				: {
						type: 'create-rectangle' as const,
						pageId: '0',
						shapeId: '1',
						x: 3,
						y: 4,
						width: 1.5,
						height: 1,
					},
			{ type: 'set-shape-data', pageId: '0', shapeId: '1', row: 'Load', data: data('Load', '60') },
			{ type: 'set-shape-data', pageId: '0', shapeId: '1', row: 'Cost', data: data('Cost', '5') },
		] as VisioEdit[])
	).bytes;
}
const page = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!;
const run = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;

describe('native data graphics in the item model', () => {
	it('applies text, bar and icon as one native graphic and reads them back', async () => {
		let bytes = await start();
		const options = { field: 'Load', min: 0, max: 100 };
		for (const kind of ['text', 'bar', 'icon'] as const) {
			const edits = visioDataGraphicEdits(await page(bytes), ['1'], { kind, ...options });
			expect(edits.map((edit) => edit.type)).toEqual(['set-data-graphic']);
			bytes = await run(bytes, edits);
		}
		const after = await page(bytes);
		expect(after.shapes[0]!.kind).toBe('group');
		expect(visioDataGraphicParts(after).size).toBe(0);
		expect(visioDataGraphicOwners(after)).toEqual(['1']);
		expect(visioDataGraphicItems(after, '1')).toEqual([
			{ kind: 'text', field: 'Load', position: 'bottom' },
			{ kind: 'bar', field: 'Load', position: 'bottom', min: 0, max: 100 },
			{ kind: 'icon', field: 'Load', position: 'top-right', min: 0, max: 100 },
		]);
		// Re-applying a kind replaces that kind.
		bytes = await run(
			bytes,
			visioDataGraphicEdits(after, ['1'], { kind: 'text', field: 'Cost', position: 'top' }),
		);
		expect(
			visioDataGraphicConfigs(await page(bytes), '1').map((c) => `${c.kind}:${c.field}`),
		).toEqual(['bar:Load', 'icon:Load', 'text:Cost']);
	});

	it('positions, auto spaces, edits and removes items', async () => {
		let bytes = await start();
		for (const field of ['Load', 'Cost'])
			bytes = await run(
				bytes,
				visioSetDataGraphicItemEdits(await page(bytes), ['1'], { kind: 'text', field }),
			);
		bytes = await run(bytes, visioPositionDataGraphicEdits(await page(bytes), ['1'], 'right'));
		let shape = (await page(bytes)).shapes[0]!;
		expect(shape.children.map((child) => child.rotation!.pinX)).toHaveLength(2);
		const ys = shape.children.map((child) => child.rotation!.pinY).sort();
		expect(ys[1]! - ys[0]!).toBeGreaterThan(0.2);
		expect(visioDataGraphicConfigs(await page(bytes), '1').map((c) => c.position)).toEqual([
			'right',
			'right',
		]);
		bytes = await run(bytes, visioAutoSpaceDataGraphicEdits(await page(bytes), ['1']));
		expect((await page(bytes)).shapes[0]!.children).toHaveLength(2);
		bytes = await run(
			bytes,
			visioSetDataGraphicItemEdits(
				await page(bytes),
				['1'],
				{ kind: 'bar', field: 'Cost', min: 0, max: 10 },
				{ kind: 'text', field: 'Load' },
			),
		);
		expect(visioDataGraphicItems(await page(bytes), '1')).toEqual([
			{ kind: 'text', field: 'Cost', position: 'right' },
			{ kind: 'bar', field: 'Cost', position: 'bottom', min: 0, max: 10 },
		]);
		bytes = await run(bytes, visioRemoveDataGraphicEdits(await page(bytes), ['1'], 'text'));
		expect(visioDataGraphicItems(await page(bytes), '1').map((item) => item.kind)).toEqual(['bar']);
		bytes = await run(bytes, visioRemoveDataGraphicEdits(await page(bytes), ['1']));
		shape = (await page(bytes)).shapes[0]!;
		expect(shape.kind).toBe('shape');
		expect(visioRemoveDataGraphicEdits(await page(bytes), ['1'])).toEqual([]);
	});

	it('moves the callouts with their shape', async () => {
		let bytes = await start();
		bytes = await run(
			bytes,
			visioDataGraphicEdits(await page(bytes), ['1'], {
				kind: 'icon',
				field: 'Load',
				min: 0,
				max: 100,
			}),
		);
		const before = (await page(bytes)).shapes[0]!;
		bytes = await run(bytes, [{ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 7 }]);
		const after = (await page(bytes)).shapes[0]!;
		// The callout is a member of the group: it keeps its place inside it and moves with it.
		expect(after.rotation!.pinX - before.rotation!.pinX).toBeCloseTo(2);
		expect(after.children[0]!.rotation).toEqual(before.children[0]!.rotation);
		expect(after.children).toHaveLength(1);
	});

	it('relays the callouts out when the shape is resized, and copies and deletes them with it', async () => {
		let bytes = await start();
		bytes = await run(
			bytes,
			visioDataGraphicEdits(await page(bytes), ['1'], { kind: 'text', field: 'Load' }),
		);
		const wide = await run(bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 2 },
		]);
		const resized = (await page(wide)).shapes[0]!;
		expect(resized.children[0]!.width).toBe(4);
		expect(resized.children[0]!.rotation!.pinY).toBeCloseTo(-0.2);
		expect(resized.children[0]!.id).toBe((await page(bytes)).shapes[0]!.children[0]!.id);
		const copied = await run(bytes, [
			{
				type: 'duplicate-shapes',
				pageId: '0',
				copies: [{ shapeId: '1', newShapeId: '10' }],
				offsetX: 1,
				offsetY: -1,
			},
		]);
		const shapes = (await page(copied)).shapes;
		expect(shapes.map((shape) => [shape.id, shape.kind, shape.children.length])).toEqual([
			['1', 'group', 1],
			['10', 'group', 1],
		]);
		const deleted = await run(copied, [{ type: 'delete-shape', pageId: '0', shapeId: '1' }]);
		expect((await page(deleted)).shapes.map((shape) => shape.id)).toEqual(['10']);
	});

	it('formats and edits the text of a shape that has a data graphic', async () => {
		let bytes = await start();
		bytes = await run(
			bytes,
			visioDataGraphicEdits(await page(bytes), ['1'], { kind: 'text', field: 'Load' }),
		);
		bytes = await run(bytes, [
			{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#5b9bd5', fillPattern: 1 },
			{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Server' },
		]);
		const shape = (await page(bytes)).shapes[0]!;
		expect(shape.style.fill).toBe('#5b9bd5');
		expect(shape.text.plainText).toBe('Server');
		expect(shape.children).toHaveLength(1);
	});

	it('keeps drawing marked shapes for an instance of a stencil master', async () => {
		const bytes = await start(true);
		const edits = visioDataGraphicEdits(await page(bytes), ['1'], { kind: 'text', field: 'Load' });
		expect(edits.map((edit) => edit.type)).not.toContain('set-data-graphic');
		const after = await run(bytes, edits);
		expect(visioDataGraphicParts(await page(after)).get('1')).toHaveLength(1);
	});
});
