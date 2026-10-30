import { describe, expect, it } from 'vitest';
import { createVsdx } from '../create-document';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import type { VisioEdit } from '../edit-commands';
import {
	visioDataGraphicItems,
	visioDeleteDataGraphicItemEdits,
	visioSetDataGraphicItemEdits,
} from './data-graphic-items';
import { visioAutoSpaceDataGraphicEdits } from './data-graphics';

const page = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!;
async function owner() {
	let bytes = (
		await editVsdx(await createVsdx(), [
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 3, y: 4, width: 1.5, height: 1 },
			{ type: 'set-shape-data', pageId: '0', shapeId: '1', row: 'Name', data: row('Name', 'Web') },
			{ type: 'set-shape-data', pageId: '0', shapeId: '1', row: 'Load', data: row('Load', '60') },
		])
	).bytes;
	bytes = (await editVsdx(bytes, [])).bytes;
	return bytes;
}
const row = (label: string, value: string) => ({
	label,
	prompt: '',
	type: 'string' as const,
	format: '',
	value,
});
const run = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;

describe('data graphic items', () => {
	it('keeps several text callouts for different fields on one shape', async () => {
		let bytes = await owner();
		for (const field of ['Name', 'Load'])
			bytes = await run(
				bytes,
				visioSetDataGraphicItemEdits(await page(bytes), ['1'], { kind: 'text', field }),
			);
		const items = visioDataGraphicItems(await page(bytes), '1');
		expect(items.map((item) => `${item.kind}:${item.field}`)).toEqual(['text:Name', 'text:Load']);
	});

	it('changes an item to another field, kind and place, and removes it', async () => {
		let bytes = await owner();
		bytes = await run(
			bytes,
			visioSetDataGraphicItemEdits(await page(bytes), ['1'], { kind: 'text', field: 'Name' }),
		);
		bytes = await run(
			bytes,
			visioSetDataGraphicItemEdits(
				await page(bytes),
				['1'],
				{ kind: 'bar', field: 'Load', position: 'top', min: 0, max: 100 },
				{ kind: 'text', field: 'Name' },
			),
		);
		expect(visioDataGraphicItems(await page(bytes), '1')).toEqual([
			{ kind: 'bar', field: 'Load', position: 'top', min: 0, max: 100 },
		]);
		bytes = await run(
			bytes,
			visioDeleteDataGraphicItemEdits(await page(bytes), ['1'], { kind: 'bar', field: 'Load' }),
		);
		expect(visioDataGraphicItems(await page(bytes), '1')).toEqual([]);
		expect((await page(bytes)).shapes.map((shape) => shape.id)).toEqual(['1']);
	});

	it('stacks items that share a place and auto space keeps both', async () => {
		let bytes = await owner();
		for (const field of ['Name', 'Load'])
			bytes = await run(
				bytes,
				visioSetDataGraphicItemEdits(await page(bytes), ['1'], {
					kind: 'text',
					field,
					position: 'right',
				}),
			);
		bytes = await run(bytes, visioAutoSpaceDataGraphicEdits(await page(bytes), ['1']));
		const after = await page(bytes);
		expect(visioDataGraphicItems(after, '1')).toHaveLength(2);
		const ys = after.shapes
			.filter((shape) => shape.id !== '1')
			.map((shape) => shape.rotation!.pinY)
			.sort();
		expect(ys[1]! - ys[0]!).toBeGreaterThan(0.2);
	});

	it('lists and recolours Color by Value as an item', async () => {
		let bytes = await owner();
		bytes = await run(
			bytes,
			visioSetDataGraphicItemEdits(await page(bytes), ['1'], { kind: 'color', field: 'Load' }),
		);
		expect(visioDataGraphicItems(await page(bytes), '1')).toEqual([
			{ kind: 'color', field: 'Load' },
		]);
		bytes = await run(
			bytes,
			visioSetDataGraphicItemEdits(
				await page(bytes),
				['1'],
				{ kind: 'color', field: 'Name' },
				{ kind: 'color', field: 'Load' },
			),
		);
		expect(visioDataGraphicItems(await page(bytes), '1')).toEqual([
			{ kind: 'color', field: 'Name' },
		]);
		bytes = await run(
			bytes,
			visioDeleteDataGraphicItemEdits(await page(bytes), ['1'], { kind: 'color', field: 'Name' }),
		);
		expect(visioDataGraphicItems(await page(bytes), '1')).toEqual([]);
	});
});
