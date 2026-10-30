import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';

/**
 * Visio 16 (COM) drew a group G (4) of a stencil Rectangle instance (1, master 2, 1.5 x 1) and a
 * plain rectangle on layer L1 (2), glued a Dynamic connector (5) from the stencil member to
 * rectangle 3, saved, then set the group's Width x1.5 and Height x2. The numbers below are what
 * it read back after the resize.
 */
const source = new Uint8Array(
	readFileSync(new URL('./__fixtures__/group-stencil-layer.vsdx', import.meta.url)),
);
const visio = {
	'1': { PinX: 1.125, PinY: 1, Width: 2.25, Height: 2, LocPinX: 1.125, LocPinY: 1 },
	'2': { PinX: 4.125, PinY: 1, Width: 1.5, Height: 2, LocPinX: 0.75, LocPinY: 1 },
};
const names = ['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY'];
/** The cached V of each transform cell per shape ID (cells precede sections and sub-shapes). */
function readShapeCells(xml: string): Record<string, Record<string, number>> {
	const result: Record<string, Record<string, number>> = {};
	for (const match of xml.matchAll(/<Shape [^>]*?ID=['"](\d+)['"][^>]*>/g)) {
		const rest = xml.slice(match.index! + match[0].length);
		const end = rest.search(/<Shapes>|<\/Shape>|<Section /);
		const cells: Record<string, number> = {};
		for (const cell of (end < 0 ? rest : rest.slice(0, end)).matchAll(
			/<Cell N=['"](\w+)['"] V=['"]([^'"]*)['"]/g,
		))
			if (names.includes(cell[1]!)) cells[cell[1]!] = Number(cell[2]);
		result[match[1]!] = cells;
	}
	return result;
}
const connector = [2.0625, 4.5, 8, 6.5];

async function resize(bytes: Uint8Array) {
	const result = await editVsdx(bytes, [
		{
			type: 'resize-shape',
			pageId: '0',
			shapeId: '4',
			width: 3.25 * 1.5,
			height: 2,
			anchor: { x: 0.5, y: 0.5 },
		},
	]);
	return (await JSZip.loadAsync(result.bytes)).file('visio/pages/page1.xml')!.async('string');
}

describe('group resize with a stencil member and a layered member (Visio 16)', () => {
	it('matches the member cells and the glued connector Visio computed', async () => {
		const xml = await resize(source);
		const actual = readShapeCells(xml);
		for (const [id, cells] of Object.entries(visio))
			for (const [name, value] of Object.entries(cells))
				expect(actual[id]![name], `shape ${id} ${name}`).toBeCloseTo(value, 6);
		const shape = /<Shape ID=['"]5['"][\s\S]*?<\/Shape>/.exec(xml)![0];
		['BeginX', 'BeginY', 'EndX', 'EndY'].forEach((name, index) =>
			expect(
				Number(new RegExp(`<Cell N=['"]${name}['"] V=['"]([^'"]*)['"]`).exec(shape)![1]),
				name,
			).toBeCloseTo(connector[index]!, 6),
		);
	});
	it('keeps the stencil member an instance with refreshed inherited caches', async () => {
		const xml = await resize(source);
		const member = /<Shape ID=['"]1['"][\s\S]*?<\/Shape>/.exec(xml)![0];
		expect(member).toContain('Master="2"');
		// The master's geometry follows the new size through its own formulas.
		expect(member).toMatch(/<Row T="LineTo" IX="2"><Cell N="X" V="2\.25" F="Inh"/);
		expect(member).toMatch(/<Cell N="PinX" V="1\.125" F="Sheet\.4!Width\*/);
	});
});
