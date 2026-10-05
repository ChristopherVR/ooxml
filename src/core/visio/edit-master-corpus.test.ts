import { DOMParser } from '@xmldom/xmldom';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit, type EditVsdxResult } from './edit.js';
import { VisioPackageError } from './package-common.js';
import { parseVsdx } from './parser.js';
import { children } from './sheet.js';

const root = process.env['VISIO_EDIT_CORPUS_DIR'];
const pagePath = 'visio/pages/page1.xml';
const hash = '5a5117a6c41313a6e7253689d7ae09c7e46bfbb958a670f60f7f5d0bbaef27f9';
const referencedHash = '558c09c518848c98a815c9684f511d5260585fa0c893f2d570b6284d6a9b29a8';
const move = (): Extract<VisioEdit, { type: 'move-shape' }> => ({
	type: 'move-shape',
	pageId: '0',
	shapeId: '1',
	x: 3,
	y: 4,
});
const resize = (): Extract<VisioEdit, { type: 'resize-shape' }> => ({
	type: 'resize-shape',
	pageId: '0',
	shapeId: '1',
	width: 4,
	height: 2,
});
const remove: VisioEdit = { type: 'delete-shape', pageId: '0', shapeId: '1' };
async function input(file = 'test_text_extraction.vsdx') {
	const bytes = new Uint8Array(await readFile(join(root!, 'poi/test-data/diagram', file)));
	expect(createHash('sha256').update(bytes).digest('hex')).toBe(
		file === 'test_text_extraction.vsdx' ? hash : referencedHash,
	);
	return bytes;
}
async function page(zip: JSZip) {
	return new DOMParser().parseFromString(await zip.file(pagePath)!.async('string'), 'text/xml')
		.documentElement! as unknown as Element;
}
function shapes(page: Element) {
	return children(children(page, 'Shapes')[0], 'Shape');
}
function target(page: Element) {
	return shapes(page).find((shape) => shape.getAttribute('ID') === '1')!;
}
function value(shape: Element, name: string) {
	return Number(
		children(shape, 'Cell')
			.find((node) => node.getAttribute('N') === name)!
			.getAttribute('V'),
	);
}
// Compare semantic XML trees, preserving child/text order and attribute values;
// serializer-added namespace declarations and prefix spellings do not change a subtree.
function canonical(node: Node): unknown {
	if (node.nodeType !== 1) return [node.nodeType, node.nodeValue];
	const element = node as Element;
	return [
		element.namespaceURI,
		element.localName,
		Array.from(element.attributes)
			.filter((attr) => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
			.map((attr) => [attr.namespaceURI, attr.localName, attr.value])
			.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
		Array.from(node.childNodes).map(canonical),
	];
}
async function preserved(original: Uint8Array, result: EditVsdxResult, deleted = false) {
	const before = await JSZip.loadAsync(original),
		after = await JSZip.loadAsync(result.bytes);
	expect(result.changedParts).toEqual([pagePath]);
	const names = (zip: JSZip) =>
		Object.values(zip.files)
			.filter((entry) => !entry.dir)
			.map((entry) => entry.name)
			.sort();
	expect(names(after)).toEqual(names(before));
	for (const [path, entry] of Object.entries(before.files)) {
		if (entry.dir || path === pagePath) continue;
		expect(await after.file(path)!.async('uint8array'), path).toEqual(
			await entry.async('uint8array'),
		);
	}
	const originalPage = await page(before),
		savedPage = await page(after);
	const siblings = (root: Element) =>
		shapes(root)
			.filter((shape) => shape.getAttribute('ID') !== '1')
			.map(canonical);
	expect(siblings(savedPage)).toEqual(siblings(originalPage));
	if (deleted) expect(target(savedPage)).toBeUndefined();
	else {
		expect(
			children(target(savedPage), 'Section')
				.filter((section) => section.getAttribute('N') === 'Geometry')
				.map(canonical),
		).toEqual(
			children(target(originalPage), 'Section')
				.filter((section) => section.getAttribute('N') === 'Geometry')
				.map(canonical),
		);
		for (const name of ['LocPinX', 'LocPinY']) {
			const cell = (root: Element) =>
				children(target(root), 'Cell').find((node) => node.getAttribute('N') === name)!;
			expect([cell(savedPage).getAttribute('F'), cell(savedPage).getAttribute('U')]).toEqual([
				cell(originalPage).getAttribute('F'),
				cell(originalPage).getAttribute('U'),
			]);
		}
	}
	const parsed = await parseVsdx(result.bytes);
	expect(parsed.pages[0]!.shapes.length).toBe(shapes(originalPage).length - Number(deleted));
	return savedPage;
}

describe.skipIf(!root)('public corpus existing local shape with preserved master parts', () => {
	it('moves the pristine text-extraction rectangle with all sibling and package data preserved', async () => {
		const bytes = await input(),
			snapshot = bytes.slice();
		const saved = await preserved(snapshot, await editVsdx(bytes, [move()]));
		expect(bytes).toEqual(snapshot);
		expect(
			['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY'].map((name) =>
				value(target(saved), name),
			),
		).toEqual([3, 4, 1.84375, 1.03125, 0.921875, 0.515625]);
	});
	it('resizes the pristine rectangle with fixed pin and recalculated local pin caches', async () => {
		const bytes = await input(),
			snapshot = bytes.slice();
		const saved = await preserved(snapshot, await editVsdx(bytes, [resize()]));
		expect(bytes).toEqual(snapshot);
		expect(
			['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY'].map((name) =>
				value(target(saved), name),
			),
		).toEqual([1.578125, 6.984375, 4, 2, 2, 1]);
	});
	it('deletes the pristine unreferenced rectangle and preserves siblings and master archive parts', async () => {
		const bytes = await input(),
			snapshot = bytes.slice();
		await preserved(snapshot, await editVsdx(bytes, [remove]), true);
		expect(bytes).toEqual(snapshot);
	});
	it('applies a combined move and resize transaction with the moved pin fixed', async () => {
		const bytes = await input(),
			snapshot = bytes.slice();
		const saved = await preserved(snapshot, await editVsdx(bytes, [move(), resize()]));
		expect(bytes).toEqual(snapshot);
		expect(
			['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY'].map((name) =>
				value(target(saved), name),
			),
		).toEqual([3, 4, 4, 2, 2, 1]);
	});
	it('owns the corpus input and mutable command values before asynchronous processing', async () => {
		const bytes = await input(),
			snapshot = bytes.slice(),
			command = move(),
			commands: VisioEdit[] = [command, resize()];
		const pending = editVsdx(bytes, commands);
		bytes.fill(0);
		command.x = 999;
		commands.length = 0;
		const saved = await preserved(snapshot, await pending);
		expect(['PinX', 'PinY', 'Width', 'Height'].map((name) => value(target(saved), name))).toEqual([
			3, 4, 4, 2,
		]);
	});
	it('refuses deletion of the explicitly referenced test.vsdx shape without mutating source', async () => {
		const bytes = await input('test.vsdx'),
			snapshot = bytes.slice();
		let error: unknown;
		try {
			await editVsdx(bytes, [remove]);
		} catch (caught) {
			error = caught;
		}
		expect(error).toBeInstanceOf(VisioPackageError);
		expect([
			'EDIT_REFERENCED_DELETE',
			'EDIT_UNKNOWN_DEPENDENCY',
			'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
		]).toContain((error as VisioPackageError).code);
		expect(bytes).toEqual(snapshot);
	});
});
