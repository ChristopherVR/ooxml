import { describe, expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { editVsdx, type VisioEdit } from './edit.js';
import { parseVsdx } from './parser.js';
import { fixture, shape } from './test-fixtures.js';
import { snapshotEdits } from './ui/edit-commands.js';
import { parseAppProperties } from '../opc/properties/index.js';

const source = () =>
	fixture({
		pages: ['0', '1', '2'].map((id) => ({ id, contents: `<Shapes>${shape('1')}</Shapes>` })),
	});
it('moves a page to the front, middle and end while preserving existing IDs, parts and relationships', async () => {
	const original = await source();
	let bytes = original;
	for (const [index, order] of [
		[0, ['2', '0', '1']],
		[1, ['0', '2', '1']],
		[2, ['0', '1', '2']],
	] as const) {
		const saved = await editVsdx(bytes, [{ type: 'reorder-page', pageId: '2', index }]);
		expect((await parseVsdx(saved.bytes)).pages.map((page) => page.id)).toEqual(order);
		expect(saved.changedParts).toEqual(['visio/pages/pages.xml']);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved.bytes);
		for (const path of Object.keys(before.files).filter(
			(path) => !before.files[path]!.dir && !saved.changedParts.includes(path),
		))
			expect(await after.file(path)!.async('uint8array')).toEqual(
				await before.file(path)!.async('uint8array'),
			);
		bytes = saved.bytes;
	}
});
it('returns unchanged bytes for an unchanged page index and snapshots page commands before awaiting', async () => {
	const original = await source();
	const noop = await editVsdx(original, [{ type: 'reorder-page', pageId: '1', index: 1 }]);
	expect(noop.bytes).toEqual(original);
	expect(noop.changedParts).toEqual([]);
	const command = { type: 'reorder-page' as const, pageId: '2', index: 0 };
	const pending = editVsdx(original, [command]);
	command.index = 2;
	expect((await parseVsdx((await pending).bytes)).pages[0]!.id).toBe('2');
	expect(snapshotEdits([{ ...command, arbitrary: 'ignored' } as VisioEdit])).toEqual([command]);
});
it.each([-1, 3, 0.5, NaN, Infinity])('rejects invalid index %s atomically', async (index) => {
	const original = await source(),
		snapshot = new Uint8Array(original);
	await expect(
		editVsdx(original, [{ type: 'reorder-page', pageId: '1', index }]),
	).rejects.toThrow();
	expect(original).toEqual(snapshot);
});
it('applies insertion and reordering in one page transaction', async () => {
	const saved = await editVsdx(await source(), [
		{ type: 'insert-page', pageId: '9', afterPageId: '2', name: 'Last' },
		{ type: 'reorder-page', pageId: '9', index: 0 },
	]);
	expect((await parseVsdx(saved.bytes)).pages.map((page) => page.id)).toEqual(['9', '0', '1', '2']);
});
const native = process.env.VISIO_NATIVE_PAGE_SCALES_DIR;
describe.skipIf(!native)('native page order', () => {
	it('matches native Visio page order and refreshes the property title order', async () => {
		const source = await readFile(resolve(native!, 'page-scales.vsdx'));
		const saved = await editVsdx(source, [{ type: 'reorder-page', pageId: '6', index: 0 }]);
		const model = await parseVsdx(saved.bytes);
		expect(model.pages.map((page) => page.id)).toEqual(['6', '0', '4', '5', '7', '8', '9']);
		expect(saved.changedParts).toEqual(['visio/pages/pages.xml', 'docProps/app.xml']);
		const zip = await JSZip.loadAsync(saved.bytes);
		expect(
			parseAppProperties(await zip.file('docProps/app.xml')!.async('string')).titlesOfParts,
		).toEqual(model.pages.map((page) => page.name));
		if (process.env.VISIO_NATIVE_PAGE_ORDER_OUTPUT)
			await writeFile(process.env.VISIO_NATIVE_PAGE_ORDER_OUTPUT, saved.bytes);
	});
});
