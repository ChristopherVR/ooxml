import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

const directory = process.env.VISIO_NATIVE_CURVE_DIR;
it.skipIf(!directory)(
	'edits and reopens native curved connectors with facing sides and corner turns',
	async () => {
		const input = new Uint8Array(readFileSync(join(directory!, 'curve-6-6-1-1.vsdx')));
		const source = await parseVsdx(input);
		const pageId = source.pages[0]!.id;
		for (const [x, y] of [
			[6, 6.2],
			[6, 6.8],
			[2.2, 3],
			[0, 6.2],
			[6, 3],
			[6, 8],
			[6, 7],
			[3, 3],
			[1, 3],
			[0, 8],
		] as const) {
			const edited = await editVsdx(input, [{ type: 'move-shape', pageId, shapeId: '2', x, y }]);
			const reopened = await parseVsdx(edited.bytes);
			expect(reopened.pages[0]!.shapes.find((shape) => shape.id === '3')!.connectorRoute).toBe(
				'curved',
			);
			writeFileSync(join(directory!, `edited-curve-${x}-${y}.vsdx`), edited.bytes);
		}
	},
);
