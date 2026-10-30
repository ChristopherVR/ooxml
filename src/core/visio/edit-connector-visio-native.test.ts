import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';

/**
 * A drawing saved by Microsoft Visio 16 (BASFLO_U Process, Decision and Dynamic connector, glued
 * with GlueTo and AutoConnect). The expected ends are what Visio itself computed when the same
 * shape was moved or the same end re-glued through COM, read back with ResultIU.
 */
const variants = () =>
	new Uint8Array(
		readFileSync(new URL('./__fixtures__/dynamic-connector-variants-native.vsdx', import.meta.url)),
	);
const pageXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string')).replace(
		/'/g,
		'"',
	);
const ends = async (bytes: Uint8Array, id: string) => {
	const shape = new RegExp(`<Shape ID="${id}".*?</Shape>`, 's').exec(await pageXml(bytes))![0];
	const cell = (name: string) =>
		Number(new RegExp(`<Cell N="${name}" V="([^"]*)"`).exec(shape)![1]);
	return [cell('BeginX'), cell('BeginY'), cell('EndX'), cell('EndY')];
};
const edit = async (...edits: VisioEdit[]) => (await editVsdx(variants(), edits)).bytes;
const move = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'move-shape',
	pageId: '0',
	shapeId,
	x,
	y,
});

describe('connectors in a drawing saved by Visio', () => {
	it('follow a moved shape to the ends Visio computes', async () => {
		const moved = await edit(move('2', 7, 7));
		expect(await ends(moved, '6')).toEqual([1.5, 8, 7.5, 7]);
		expect(await ends(moved, '7')).toEqual([7.5, 7, 10, 7.625]);
		expect(await ends(moved, '9')).toEqual([2, 4.375, 6.5, 7]);
		// A connector that does not touch the shape stays as saved.
		expect(await ends(moved, '8')).toEqual([2.5, 8, 5.5, 4]);
	});

	it('follow a shape glued through AutoConnect (whole-shape glue named PinY)', async () => {
		const moved = await edit(move('4', 3, 2));
		expect(await ends(moved, '9')).toEqual([3, 2.375, 5.5, 8]);
		expect(await ends(moved, '10')).toEqual([2, 7.625, 2.5, 2]);
		const first = await edit(move('1', 3, 9.5));
		expect(await ends(first, '6')).toEqual([2.5, 9.5, 6.5, 8]);
		expect(await ends(first, '10')).toEqual([3, 9.125, 2.5, 4]);
	});

	it('re-glues an end of a PinY-glued connector as Visio does', async () => {
		const bytes = await edit({
			type: 'glue-connector',
			pageId: '0',
			shapeId: '10',
			endpoint: 'end',
			target: '5',
		});
		expect(await ends(bytes, '10')).toEqual([2.5, 8, 5.5, 4]);
		expect(await pageXml(bytes)).toMatch(/FromSheet="10" FromCell="EndX"[^>]*ToSheet="5"/);
	});

	it('drags an end free and keeps the other end glued', async () => {
		const bytes = await edit({
			type: 'move-line-endpoint',
			pageId: '0',
			shapeId: '10',
			endpoint: 'end',
			x: 8,
			y: 2,
		});
		expect(await ends(bytes, '10')).toEqual([2.5, 8, 8, 2]);
	});
});
