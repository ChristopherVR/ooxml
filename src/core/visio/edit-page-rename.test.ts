import { expect, it, describe } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { editVsdx, type VisioPageRename } from './edit.js';
import { parseVsdx } from './parser.js';
import { cell, fixture, shape } from './test-fixtures.js';
import evidence from './__fixtures__/page-rename-native.json';

const rename: VisioPageRename = { type: 'rename-page', pageId: '0', name: 'Renamed & Page' };
const fields = (extra = '') =>
	`<Section N="User"><Row N="Name"><Cell N="Value" V="Page 1" U="STR" F="PAGENAME()"/></Row><Row N="Universal"><Cell N="Value" V="Page 1" U="STR" F="PAGENAME(750)"/></Row>${extra}</Section>`;
const source = (extra = '') =>
	fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', fields(extra))}</Shapes>` },
			{
				id: '1',
				contents: `<Shapes>${shape('1', `<Section N="User"><Row N="Width">${cell('Value', 8.5, 'Pages[Page 1]!ThePage!PageWidth')}</Row><Row N="Literal"><Cell N="Value" V="Pages[Page 1]!" U="STR" F="&quot;Pages[Page 1]!&quot;"/></Row></Section><Section N="Hyperlink"><Row IX="0"><Cell N="SubAddress" V="Page 1" F="Pages[Page 1]!ThePage!PAGENAME()"/></Row><Row IX="1"><Cell N="SubAddress" V="Page 1/Sheet.1" F="&quot;Page 1/Sheet.1&quot;"/></Row></Section>`)}</Shapes>`,
			},
		],
	});
const xml = async (bytes: Uint8Array, path: string) =>
	(await JSZip.loadAsync(bytes)).file(path)!.async('string');
it('renames the first local/universal name, static references and direct string caches', async () => {
	const saved = await editVsdx(await source(), [rename]);
	expect((await parseVsdx(saved.bytes)).pages.map((page) => [page.id, page.name])).toEqual([
		['0', rename.name],
		['1', 'Page 2'],
	]);
	expect(await xml(saved.bytes, 'visio/pages/pages.xml')).toContain('NameU="Renamed &amp; Page"');
	expect(await xml(saved.bytes, 'visio/pages/page1.xml')).toContain(
		'V="Renamed &amp; Page" U="STR" F="PAGENAME()"',
	);
	const other = await xml(saved.bytes, 'visio/pages/page2.xml');
	expect(other).toContain('F="Pages[Renamed &amp; Page]!ThePage!PageWidth"');
	expect(other).toContain(
		'V="Renamed &amp; Page" F="Pages[Renamed &amp; Page]!ThePage!PAGENAME()"',
	);
	expect(other).toContain('V="Renamed &amp; Page/Sheet.1"');
	expect(other).toContain('F="&quot;Pages[Page 1]!&quot;"');
});
it('preserves a custom universal name on subsequent local renames', async () => {
	const first = await editVsdx(await source(), [rename]);
	const saved = await editVsdx(first.bytes, [{ ...rename, name: 'Local only' }]);
	expect(await xml(saved.bytes, 'visio/pages/pages.xml')).toContain('NameU="Renamed &amp; Page"');
	const content = await xml(saved.bytes, 'visio/pages/page1.xml');
	expect(content).toContain('V="Local only" U="STR" F="PAGENAME()"');
	expect(content).toContain('V="Renamed &amp; Page" U="STR" F="PAGENAME(750)"');
	expect(await xml(saved.bytes, 'visio/pages/page2.xml')).toContain(
		'F="Pages[Renamed &amp; Page]!ThePage!PageWidth"',
	);
});
it('matches recorded native local/universal rename semantics', async () => {
	const first = await editVsdx(await source(), [{ ...rename, name: evidence.first.name }]);
	const second = await editVsdx(first.bytes, [{ ...rename, name: evidence.second.name }]);
	for (const [bytes, expected] of [
		[first.bytes, evidence.first],
		[second.bytes, evidence.second],
	] as const) {
		expect((await parseVsdx(bytes)).pages[0]!.name).toBe(expected.name);
		expect(await xml(bytes, 'visio/pages/pages.xml')).toContain(
			`NameU="${expected.nameU.replaceAll('&', '&amp;')}"`,
		);
		expect(evidence.accepted).toContainEqual(expected);
	}
});
it.each(['page 2', '', 'a\nb', 'x'.repeat(256)])(
	'rejects invalid or duplicate names atomically: %j',
	async (name) => {
		const original = await source(),
			snapshot = new Uint8Array(original);
		await expect(editVsdx(original, [{ ...rename, name }])).rejects.toThrow();
		expect(original).toEqual(snapshot);
	},
);
it.each([
	'<Row N="Derived"><Cell N="Value" V="Page 1" U="STR" F="User.Name"/></Row>',
	'<Row N="Derived"><Cell N="Value" V="Page 1!" U="STR" F="CAT(PAGENAME(),&quot;!&quot;)"/></Row>',
	'<Row N="Derived"><Cell N="Value" V="Page 1" U="STR" F="EVALCELL(&quot;User.Name&quot;)"/></Row>',
])('refuses unsupported string-valued page-name dependency closure', async (extra) => {
	await expect(editVsdx(await source(extra), [rename])).rejects.toThrow(/page-name/);
});
it('preserves unrelated payload bytes and leaves no-op bytes exact', async () => {
	const original = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1')}</Shapes>` }],
	});
	const unchanged = await editVsdx(original, [{ ...rename, name: 'Page 1' }]);
	expect(unchanged.bytes).toEqual(original);
	const saved = await editVsdx(original, [rename]);
	expect(saved.changedParts).toEqual(['visio/pages/pages.xml']);
	expect(await xml(saved.bytes, 'visio/pages/page1.xml')).toBe(
		await xml(original, 'visio/pages/page1.xml'),
	);
});
it('snapshots rename input before the first asynchronous operation', async () => {
	const command = { ...rename },
		original = await source();
	const pending = editVsdx(original, [command]);
	command.name = 'Mutated';
	expect((await parseVsdx((await pending).bytes)).pages[0]!.name).toBe(rename.name);
});
it('follows page relationships outside the conventional folder and preserves opaque XML', async () => {
	const original = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', fields())}</Shapes>` }],
		edit: (zip) => {
			zip.file('custom/page.xml', zip.file('visio/pages/page1.xml')!.async('string'));
			zip.remove('visio/pages/page1.xml');
			zip.file(
				'visio/pages/_rels/pages.xml.rels',
				zip
					.file('visio/pages/_rels/pages.xml.rels')!
					.async('string')
					.then((xml) => xml.replace('page1.xml', '../../custom/page.xml')),
			);
			zip.file('visio/opaque.xml', '<opaque broken');
		},
	});
	const saved = await editVsdx(original, [rename]);
	expect(await xml(saved.bytes, 'custom/page.xml')).toContain('V="Renamed &amp; Page"');
	expect(await xml(saved.bytes, 'visio/opaque.xml')).toBe('<opaque broken');
});

const native = process.env.VISIO_NATIVE_PAGE_RENAME_DIR;
describe.skipIf(!native)('native page rename', () => {
	it('produces native-equivalent names, reference formulas and direct caches', async () => {
		const saved = await editVsdx(await readFile(resolve(native!, 'original.vsdx')), [rename]);
		await writeFile(resolve(native!, 'core-renamed.vsdx'), saved.bytes);
		const second = await editVsdx(saved.bytes, [{ ...rename, name: 'Second rename' }]);
		await writeFile(resolve(native!, 'core-renamed-again.vsdx'), second.bytes);
		for (const path of ['visio/pages/page1.xml', 'visio/pages/page2.xml']) {
			const cells = (source: string) =>
				[
					...source.matchAll(
						/<Cell N=['"](?:Value|SubAddress)['"] V=['"]([^'"]+)['"][^>]*F=['"]([^'"]+)['"]/g,
					),
				].map((match) => [match[1], match[2]]);
			const nativeBytes = await readFile(resolve(native!, 'renamed-native.vsdx'));
			const reference = cells(await xml(nativeBytes, path));
			expect(reference.length).toBeGreaterThan(0);
			expect(cells(await xml(saved.bytes, path))).toEqual(reference);
		}
	});
});
