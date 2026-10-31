import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { parseVsdx } from './parser';
import { cell, fixture, row, section, shape } from './test-fixtures';

const size =
	cell('Width', 1) +
	cell('Height', 1) +
	cell('PinX', 0) +
	cell('PinY', 0) +
	cell('LocPinX', 0) +
	cell('LocPinY', 0);
const line = (x = 1, y = 1) =>
	section(
		'Geometry',
		row(1, 'InfiniteLine', cell('X', 0) + cell('Y', 0) + cell('A', x) + cell('B', y)),
	);

describe('infinite-line master previews', () => {
	it('draws the three line orientations saved by Visio 16', async () => {
		const bytes = readFileSync(
			new URL('./__fixtures__/infinite-line-masters-native.vsdx', import.meta.url),
		);
		const model = await parseVsdx(new Uint8Array(bytes));
		expect(
			model.masters!.map((master) => [
				master.name,
				master.width,
				master.height,
				master.shapes[0]!.geometry,
			]),
		).toEqual([
			['Diagonal', 4, 3, [{ path: 'M 0 0 L 3 3', fill: false, stroke: true }]],
			['Horizontal', 4, 3, [{ path: 'M 0 0.5 L 4 0.5', fill: false, stroke: true }]],
			['Vertical', 4, 3, [{ path: 'M 0.5 0 L 0.5 3', fill: false, stroke: true }]],
		]);
	});

	it('clips each master to its own page and does not add drawing diagnostics', async () => {
		const model = await parseVsdx(
			await fixture({
				masters: [
					{
						id: '1',
						sheet: cell('PageWidth', 10) + cell('PageHeight', 6),
						shapes: shape('1', size + line()),
					},
					{
						id: '2',
						sheet: cell('PageWidth', 4) + cell('PageHeight', 3),
						shapes: shape('2', size + line()),
					},
				],
			}),
		);
		expect(model.masters!.map((master) => master.shapes[0]!.geometry)).toEqual([
			[{ path: 'M 0 0 L 6 6', fill: false, stroke: true }],
			[{ path: 'M 0 0 L 3 3', fill: false, stroke: true }],
		]);
		expect(model.diagnostics).toEqual((await parseVsdx(await fixture())).diagnostics);
	});

	it('uses a local box when the next master omits or has invalid page dimensions', async () => {
		const model = await parseVsdx(
			await fixture({
				masters: [
					{
						id: '1',
						sheet: cell('PageWidth', 10) + cell('PageHeight', 6),
						shapes: shape('1', size + line()),
					},
					{ id: '2', shapes: shape('2', size + cell('PinX', 20) + line()) },
					{
						id: '3',
						sheet: cell('PageWidth', -1) + cell('PageHeight', 6),
						shapes: shape('3', size + line()),
					},
				],
			}),
		);
		for (const master of model.masters!.slice(1))
			expect(master.shapes[0]!.geometry).toEqual([
				{ path: 'M 0 0 L 1 1', fill: false, stroke: true },
			]);
	});

	it('clips a nested line through its rotated and translated ancestor', async () => {
		const model = await parseVsdx(
			await fixture({
				masters: [
					{
						id: '1',
						sheet: cell('PageWidth', 4) + cell('PageHeight', 3),
						shapes: shape(
							'1',
							size +
								cell('PinX', 2) +
								cell('PinY', 1) +
								cell('Angle', Math.PI / 2) +
								`<Shapes>${shape('2', size + line(1, 0))}</Shapes>`,
							'Type="Group"',
						),
					},
				],
			}),
		);
		expect(model.masters![0]!.shapes[0]!.children[0]!.geometry).toEqual([
			{ path: 'M -1 0 L 2 0', fill: false, stroke: true },
		]);
	});
});
