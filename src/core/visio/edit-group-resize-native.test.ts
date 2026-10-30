import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import evidence from './__fixtures__/group-nested-rotated.json';

/**
 * Visio 16 built this drawing through COM (a group holding a rectangle, a 90 degree rectangle and a
 * nested group of a 30 degree rectangle and a FlipY rectangle), saved it, then set the outer group's
 * Width x1.5 and Height x2. `evidence.after` is every member cell read back from Visio. Rotated
 * members scale by W(cos^2 sx + sin^2 sy) and H(sin^2 sx + cos^2 sy); pins scale by sx and sy.
 */
const source = readFileSync(new URL('./__fixtures__/group-nested-rotated.vsdx', import.meta.url));
const names = ['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY', 'Angle', 'FlipX', 'FlipY'];

/** The cached V of each transform cell, per shape ID (cells precede the nested Shapes element). */
export function readShapeCells(xml: string): Record<string, Record<string, number>> {
	const result: Record<string, Record<string, number>> = {};
	for (const match of xml.matchAll(/<Shape [^>]*?ID=['"](\d+)['"][^>]*>/g)) {
		const rest = xml.slice(match.index! + match[0].length);
		const end = rest.search(/<Shapes>|<Shapes |<\/Shape>|<Section /);
		const body = end < 0 ? rest : rest.slice(0, end);
		const cells: Record<string, number> = {};
		for (const cell of body.matchAll(/<Cell N=['"](\w+)['"] V=['"]([^'"]*)['"]/g))
			if (names.includes(cell[1]!)) cells[cell[1]!] = Number(cell[2]);
		result[match[1]!] = cells;
	}
	return result;
}

describe('group resize against Visio 16 (nested, rotated and flipped members)', () => {
	it('matches every member cell Visio computed', async () => {
		const { outer, factor, before, after } = evidence;
		const group = before[outer as keyof typeof before] as Record<string, number>;
		const result = await editVsdx(new Uint8Array(source), [
			{
				type: 'resize-shape',
				pageId: '0',
				shapeId: outer,
				width: group.Width! * factor.x,
				height: group.Height! * factor.y,
				anchor: { x: 0.5, y: 0.5 },
			},
		]);
		const page = await (
			await JSZip.loadAsync(result.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		const actual = readShapeCells(page);
		for (const [id, expected] of Object.entries(after)) {
			if (id === outer) continue;
			for (const name of names)
				expect(actual[id]![name], `shape ${id} ${name}`).toBeCloseTo(
					(expected as Record<string, number>)[name]!,
					6,
				);
		}
	});
});
