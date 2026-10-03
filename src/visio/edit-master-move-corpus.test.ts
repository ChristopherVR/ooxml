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
const command = () => ({ type: 'move-shape' as const, pageId: '0', shapeId: '3', x: 6, y: 7 });
async function source() {
	const bytes = new Uint8Array(
		await readFile(join(directory!, 'poi/test-data/diagram/test_text_extraction.vsdx')),
	);
	expect(createHash('sha256').update(bytes).digest('hex')).toBe(
		'5a5117a6c41313a6e7253689d7ae09c7e46bfbb958a670f60f7f5d0bbaef27f9',
	);
	return bytes;
}
async function verify(beforeBytes: Uint8Array, afterBytes: Uint8Array) {
	const before = await JSZip.loadAsync(beforeBytes),
		after = await JSZip.loadAsync(afterBytes);
	const path = 'visio/pages/page1.xml';
	for (const [name, part] of Object.entries(before.files))
		if (!part.dir && name !== path)
			expect(await after.file(name)!.async('uint8array'), name).toEqual(
				await part.async('uint8array'),
			);
	const oldPage = parseXml(await before.file(path)!.async('string')).documentElement;
	const newPage = parseXml(await after.file(path)!.async('string')).documentElement;
	const shapes = (root: Element) => children(children(root, 'Shapes')[0], 'Shape');
	const target = (root: Element) => shapes(root).find((shape) => shape.getAttribute('ID') === '3')!;
	const saved = target(newPage),
		original = target(oldPage);
	expect(saved.getAttribute('Master')).toBe('4');
	const pinned = (node: Element) =>
		children(node, 'Cell').filter((cell) => ['PinX', 'PinY'].includes(cell.getAttribute('N')!));
	const pins = pinned(saved).map((cell) => Number(cell.getAttribute('V')));
	for (const root of [oldPage, newPage])
		for (const cell of pinned(target(root))) cell.parentNode!.removeChild(cell);
	// Normalize both documents through the same serializer before comparing owned subtrees.
	const normalizedOld = parseXml(buildXml(oldPage.ownerDocument!)).documentElement;
	const normalizedNew = parseXml(buildXml(newPage.ownerDocument!)).documentElement;
	const serialize = (node: Element) =>
		new XMLSerializer().serializeToString(node as unknown as XmlNode);
	expect(shapes(normalizedNew).map(serialize)).toEqual(shapes(normalizedOld).map(serialize));
	return pins;
}
describe.skipIf(!directory)('immutable public master-linked Pentagon move', () => {
	it('moves and reverses only the pristine local pin overrides while preserving all inheritance', async () => {
		const bytes = await source(),
			snapshot = bytes.slice();
		const before = await parseVsdx(bytes),
			original = before.pages[0]!.shapes.find((shape) => shape.id === '3')!;
		const moved = await editVsdx(bytes, [command()]);
		expect(moved.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(await verify(snapshot, moved.bytes)).toEqual([6, 7]);
		expect(bytes).toEqual(snapshot);
		const after = (await parseVsdx(moved.bytes)).pages[0]!.shapes.find(
			(shape) => shape.id === '3',
		)!;
		expect([after.width, after.height]).toEqual([original.width, original.height]);
		expect(after.geometry).toEqual(original.geometry);
		expect(after.transform[4]).toBeCloseTo(original.transform[4] + 6 - 5.171428571428571, 12);
		expect(after.transform[5]).toBeCloseTo(original.transform[5] + 7 - 2.34164075203105, 12);
		const reversed = await editVsdx(moved.bytes, [
			{ ...command(), x: 5.171428571428571, y: 2.34164075203105 },
		]);
		expect(await verify(snapshot, reversed.bytes)).toEqual([5.171428571428571, 2.34164075203105]);
		expect(
			(await parseVsdx(reversed.bytes)).pages[0]!.shapes.find((shape) => shape.id === '3')!
				.transform,
		).toEqual(original.transform);
	});
	it('owns pristine public bytes and commands before asynchronous processing', async () => {
		const bytes = await source(),
			snapshot = bytes.slice(),
			edit = command();
		const pending = editVsdx(bytes, [edit]);
		bytes.fill(0);
		edit.x = 999;
		edit.shapeId = 'missing';
		expect(await verify(snapshot, (await pending).bytes)).toEqual([6, 7]);
	});
});
