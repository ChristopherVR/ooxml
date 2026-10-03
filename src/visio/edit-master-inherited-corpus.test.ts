import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { XMLSerializer, type Node as XmlNode } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import { parseXml, buildXml } from '../xml/index.js';
import { editVsdx } from './edit.js';
import { parseVsdx } from './parser.js';
import { children } from './sheet.js';
const directory = process.env['VISIO_EDIT_CORPUS_DIR'];
const targets = [
	[
		'libvisio/src/test/data/bgcolor.vsdx',
		'1',
		'7b89f56666ec41b4e1122d3ffcaa11ec1c102ec42c5a11b1a95617c4626c8ae0',
	],
	[
		'libvisio/src/test/data/bgcolor.vsdx',
		'2',
		'7b89f56666ec41b4e1122d3ffcaa11ec1c102ec42c5a11b1a95617c4626c8ae0',
	],
	[
		'libvisio/src/test/data/bgcolor.vsdx',
		'3',
		'7b89f56666ec41b4e1122d3ffcaa11ec1c102ec42c5a11b1a95617c4626c8ae0',
	],
	[
		'libvisio/src/test/data/dwg.vsdx',
		'2',
		'e05d55753bdc206dac427ec50e95fcc5672b2067e6ce9384682dfc508af5e189',
	],
	[
		'libvisio/src/test/data/dwg.vsdx',
		'3',
		'e05d55753bdc206dac427ec50e95fcc5672b2067e6ce9384682dfc508af5e189',
	],
	[
		'libvisio/src/test/data/dwg.vsdx',
		'4',
		'e05d55753bdc206dac427ec50e95fcc5672b2067e6ce9384682dfc508af5e189',
	],
	[
		'libvisio/src/test/data/fdo86664.vsdx',
		'1',
		'fc628bbb1c4686489a6b5bf7c2f9987db870ac2bb12dd8a6c9c27ed501a338b5',
	],
	[
		'libvisio/src/test/data/fdo86664.vsdx',
		'3',
		'fc628bbb1c4686489a6b5bf7c2f9987db870ac2bb12dd8a6c9c27ed501a338b5',
	],
	[
		'poi/test-data/diagram/test_text_extraction.vsdx',
		'2',
		'5a5117a6c41313a6e7253689d7ae09c7e46bfbb958a670f60f7f5d0bbaef27f9',
	],
	[
		'libvisio/src/test/data/dwg.vsdx',
		'1',
		'e05d55753bdc206dac427ec50e95fcc5672b2067e6ce9384682dfc508af5e189',
	],
	[
		'libvisio/src/test/data/fdo86664.vsdx',
		'2',
		'fc628bbb1c4686489a6b5bf7c2f9987db870ac2bb12dd8a6c9c27ed501a338b5',
	],
] as const;
async function preserved(original: Uint8Array, saved: Uint8Array, id: string) {
	const before = await JSZip.loadAsync(original),
		after = await JSZip.loadAsync(saved);
	for (const [name, part] of Object.entries(before.files))
		if (!part.dir && name !== 'visio/pages/page1.xml')
			expect(await after.file(name)!.async('uint8array'), name).toEqual(
				await part.async('uint8array'),
			);
	const roots = [before, after].map(
		async (zip) =>
			parseXml(await zip.file('visio/pages/page1.xml')!.async('string')).documentElement,
	);
	const normalized = await Promise.all(
		roots.map(async (promise) => {
			const root = await promise,
				shape = children(children(root, 'Shapes')[0], 'Shape').find(
					(node) => node.getAttribute('ID') === id,
				)!;
			for (const cell of children(shape, 'Cell').filter((node) =>
				['PinX', 'PinY'].includes(node.getAttribute('N')!),
			))
				shape.removeChild(cell);
			return parseXml(buildXml(root.ownerDocument!)).documentElement;
		}),
	);
	const serializer = new XMLSerializer();
	expect(
		normalized.map((root) =>
			children(children(root, 'Shapes')[0], 'Shape').map((node) =>
				serializer.serializeToString(node as unknown as XmlNode),
			),
		)[1],
	).toEqual(
		normalized.map((root) =>
			children(children(root, 'Shapes')[0], 'Shape').map((node) =>
				serializer.serializeToString(node as unknown as XmlNode),
			),
		)[0],
	);
}
describe.skipIf(!directory)('immutable public inherited master transform moves', () => {
	it.each(targets)('moves and reverses %s shape %s', async (file, id, hash) => {
		const bytes = new Uint8Array(await readFile(join(directory!, file))),
			snapshot = bytes.slice();
		expect(createHash('sha256').update(bytes).digest('hex')).toBe(hash);
		const original = (await parseVsdx(bytes)).pages[0]!.shapes.find((shape) => shape.id === id)!;
		const zip = await JSZip.loadAsync(bytes),
			page = parseXml(await zip.file('visio/pages/page1.xml')!.async('string')).documentElement;
		const target = children(children(page, 'Shapes')[0], 'Shape').find(
			(node) => node.getAttribute('ID') === id,
		)!;
		const value = (name: string) =>
			Number(
				children(target, 'Cell')
					.find((node) => node.getAttribute('N') === name)!
					.getAttribute('V'),
			);
		const x = value('PinX'),
			y = value('PinY');
		const command = {
			type: 'move-shape' as const,
			pageId: '0',
			shapeId: id,
			x: x + 0.25,
			y: y + 0.5,
		};
		const result = await editVsdx(bytes, [command]);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
		await preserved(snapshot, result.bytes, id);
		const moved = (await parseVsdx(result.bytes)).pages[0]!.shapes.find(
			(shape) => shape.id === id,
		)!;
		expect([moved.width, moved.height, moved.geometry]).toEqual([
			original.width,
			original.height,
			original.geometry,
		]);
		expect(moved.transform.slice(0, 4)).toEqual(original.transform.slice(0, 4));
		expect(moved.transform[4]).toBeCloseTo(original.transform[4] + 0.25, 12);
		expect(moved.transform[5]).toBeCloseTo(original.transform[5] + 0.5, 12);
		const reversed = await editVsdx(result.bytes, [{ ...command, x, y }]);
		await preserved(snapshot, reversed.bytes, id);
		expect(
			(await parseVsdx(reversed.bytes)).pages[0]!.shapes.find((shape) => shape.id === id)!
				.transform,
		).toEqual(original.transform);
		expect(bytes).toEqual(snapshot);
	});
	it('owns inherited Router source bytes and command values before yielding', async () => {
		const snapshot = new Uint8Array(await readFile(join(directory!, targets[8][0]))),
			bytes = snapshot.slice();
		expect(createHash('sha256').update(snapshot).digest('hex')).toBe(targets[8][2]);
		const command = { type: 'move-shape' as const, pageId: '0', shapeId: '2', x: 6, y: 7 };
		const pending = editVsdx(bytes, [command]);
		bytes.fill(0);
		command.x = 999;
		command.shapeId = '999';
		const result = await pending;
		await preserved(snapshot, result.bytes, '2');
		const shape = (await parseVsdx(result.bytes)).pages[0]!.shapes.find((node) => node.id === '2')!;
		expect(shape.transform[4]).toBeCloseTo(5.25, 12);
		expect(shape.transform[5]).toBeCloseTo(7 - 0.433333333333329, 12);
	});
	it('keeps a pristine inherited Router move and unsupported resize transaction atomic', async () => {
		const bytes = new Uint8Array(await readFile(join(directory!, targets[8][0]))),
			snapshot = bytes.slice();
		await expect(
			editVsdx(bytes, [
				{ type: 'move-shape', pageId: '0', shapeId: '2', x: 6, y: 7 },
				{ type: 'resize-shape', pageId: '0', shapeId: '2', width: 4, height: 4 },
			]),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_GEOMETRY_EDIT' });
		expect(bytes).toEqual(snapshot);
	});
});
