import { describe, expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { fixture, shape, cell } from './test-fixtures';
import evidence from './__fixtures__/page-delete-native.json';
const remove = { type: 'delete-page', pageId: '0' } as const;
const source = () =>
	fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1')}</Shapes>` },
			{
				id: '4',
				contents: `<Shapes>${shape('1', `<Section N="User"><Row N="Width">${cell('Value', 8.5, 'Pages[Page 1]!ThePage!PageWidth')}</Row></Section><Section N="Hyperlink"><Row IX="0"><Cell N="SubAddress" V="Page 1" F="Pages[Page 1]!ThePage!PAGENAME()"/></Row></Section>`)}</Shapes>`,
			},
		],
	});
const xml = async (bytes: Uint8Array, path: string) =>
	(await JSZip.loadAsync(bytes)).file(path)!.async('string');
it('deletes the page and freezes cached static references without renaming remaining pages', async () => {
	const original = await source(),
		snapshot = new Uint8Array(original);
	const saved = await editVsdx(original, [remove]);
	expect(original).toEqual(snapshot);
	expect((await parseVsdx(saved.bytes)).pages.map((page) => [page.id, page.name])).toEqual([
		['4', 'Page 2'],
	]);
	expect((await JSZip.loadAsync(saved.bytes)).file('visio/pages/page1.xml')).toBeNull();
	const remaining = await xml(saved.bytes, 'visio/pages/page2.xml');
	expect(remaining).toContain('<Cell N="Value" V="8.5"/>');
	expect(remaining).toContain('<Cell N="SubAddress" V="Page 1"/>');
	expect(saved.changedParts).toContain('visio/pages/page1.xml');
});
it('removes page-owned relationships and overrides but preserves shared and opaque assets', async () => {
	const original = await fixture({
		pages: [
			{ id: '0', contents: '<Shapes/>' },
			{ id: '1', contents: '<Shapes/>' },
		],
		edit: (zip) => {
			zip.file(
				'visio/pages/_rels/page1.xml.rels',
				'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="urn:asset" Target="../../custom/shared.bin"/></Relationships>',
			);
			zip.file('custom/shared.bin', new Uint8Array([0, 1, 255]));
			zip.file('visio/opaque.xml', '<unparsed opaque');
		},
	});
	const saved = await editVsdx(original, [remove]);
	const zip = await JSZip.loadAsync(saved.bytes);
	expect(zip.file('visio/pages/_rels/page1.xml.rels')).toBeNull();
	expect(await zip.file('custom/shared.bin')!.async('uint8array')).toEqual(
		new Uint8Array([0, 1, 255]),
	);
	expect(await xml(saved.bytes, 'visio/opaque.xml')).toBe('<unparsed opaque');
});
it('clears background references when deleting a background page', async () => {
	const original = await fixture({
		pages: [
			{ id: '0', attributes: 'BackPage="4"', contents: '<Shapes/>' },
			{ id: '4', attributes: 'Background="1"', contents: '<Shapes/>' },
		],
	});
	const saved = await editVsdx(original, [{ ...remove, pageId: '4' }]);
	expect(await xml(saved.bytes, 'visio/pages/pages.xml')).not.toContain('BackPage');
	expect((await parseVsdx(saved.bytes)).pages).toHaveLength(1);
});
it('replaces the last foreground page with a new blank page and default physical dimensions', async () => {
	const original = await fixture({
		pages: [
			{
				id: '0',
				width: 10,
				height: 12,
				pageCells: cell('DrawingScale', 2),
				contents: `<Shapes>${shape('1')}</Shapes>`,
			},
		],
	});
	const saved = await editVsdx(original, [remove]);
	const page = (await parseVsdx(saved.bytes)).pages[0]!;
	expect(page).toMatchObject({
		id: '1',
		name: 'Page-2',
		width: 8.5,
		height: 11,
		shapes: [],
		isBackground: false,
	});
	expect(page.drawingToPageScale).toBeUndefined();
	const reference = evidence.cases.find((item) => item.name === 'last')!.native.pages[0]!;
	expect([page.width, page.height, page.name, page.shapes.length]).toEqual([
		reference.width,
		reference.height,
		reference.name,
		reference.shapes,
	]);
});
it('keeps a foreground page when the deleted foreground page leaves only backgrounds', async () => {
	const original = await fixture({
		pages: [
			{ id: '0', attributes: 'BackPage="4"', contents: '<Shapes/>' },
			{ id: '4', attributes: 'Background="1"', contents: '<Shapes/>' },
		],
	});
	const saved = await editVsdx(original, [remove]);
	expect((await parseVsdx(saved.bytes)).pages.map((page) => [page.id, page.isBackground])).toEqual([
		['4', true],
		['5', false],
	]);
});
it('refreshes foreground counts, numbering and transitive caches after deletion', async () => {
	const fields =
		'<Section N="User"><Row N="Number">' +
		cell('Value', 2, 'PAGENUMBER()') +
		'</Row><Row N="Count">' +
		cell('Value', 2, 'PAGECOUNT()') +
		'</Row><Row N="Sum">' +
		cell('Value', 4, 'User.Number+User.Count') +
		'</Row></Section>';
	const original = await fixture({
		pages: [
			{ id: '0', contents: '<Shapes/>' },
			{ id: '1', contents: `<Shapes>${shape('1', fields)}</Shapes>` },
		],
	});
	const saved = await editVsdx(original, [remove]);
	const result = await xml(saved.bytes, 'visio/pages/page2.xml');
	expect(result).toContain('V="1" F="PAGENUMBER()"');
	expect(result).toContain('V="1" F="PAGECOUNT()"');
	expect(result).toContain('V="2" F="User.Number+User.Count"');
});
it('supports insertion and deletion in one ordered page transaction', async () => {
	const saved = await editVsdx(await source(), [
		{ type: 'insert-page', pageId: '8', afterPageId: '0', name: 'New' },
		remove,
	]);
	expect((await parseVsdx(saved.bytes)).pages.map((page) => page.id)).toEqual(['8', '4']);
});
it('rejects missing targets and unsupported dynamic dependencies atomically', async () => {
	const original = await source(),
		snapshot = new Uint8Array(original);
	await expect(editVsdx(original, [{ ...remove, pageId: 'missing' }])).rejects.toThrow(
		/does not exist/,
	);
	const dynamic = await fixture({
		pages: [
			{ id: '0', contents: '<Shapes/>' },
			{
				id: '1',
				contents: `<Shapes>${shape('1', cell('Width', 1, 'EVALCELL(&quot;Pages[Page 1]!ThePage!PageWidth&quot;)'))}</Shapes>`,
			},
		],
	});
	await expect(editVsdx(dynamic, [remove])).rejects.toThrow(/Dynamic page deletion/);
	expect(original).toEqual(snapshot);
});

const native = process.env.VISIO_NATIVE_PAGE_DELETE_DIR;
it('rejects unsupported explicit XML references rather than leaving dangling relationship IDs', async () => {
	const original = await fixture({
		pages: [
			{ id: '0', contents: '<Shapes/>' },
			{ id: '1', contents: '<Rel r:id="page"/>' },
		],
		edit: (zip) => {
			zip.file(
				'visio/pages/_rels/page2.xml.rels',
				'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="page" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page1.xml"/></Relationships>',
			);
		},
	});
	await expect(editVsdx(original, [remove])).rejects.toThrow(/unsupported explicit XML/);
});
describe.skipIf(!native)('native page deletion', () => {
	it('deletes native referenced, background and last-page drawings', async () => {
		for (const [source, name, id] of [
			['referenced-original.vsdx', 'referenced', '0'],
			['background-original.vsdx', 'background', '4'],
			['last-original.vsdx', 'last', '0'],
		]) {
			const saved = await editVsdx(await readFile(resolve(native!, source!)), [
				{ ...remove, pageId: id! },
			]);
			await writeFile(resolve(native!, `core-${name}.vsdx`), saved.bytes);
			expect(
				(await parseVsdx(saved.bytes)).pages.filter((page) => !page.isBackground),
			).toHaveLength(1);
		}
	});
});
