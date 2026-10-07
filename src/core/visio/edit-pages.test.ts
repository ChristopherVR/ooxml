import { describe, expect, it } from 'vitest';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { editVsdx } from './edit.js';
import { parseVsdx } from './parser.js';
import { cell, fixture, shape } from './test-fixtures.js';
import type { VisioPageInsert } from './edit-commands.js';
import { parseAppProperties } from '../opc/properties/index.js';
import nativeEvidence from './__fixtures__/page-insertion-native.json';

const insert: VisioPageInsert = {
	type: 'insert-page',
	pageId: '99',
	afterPageId: '0',
	name: 'Inserted & Page',
};
const source = () =>
	fixture({
		pages: [
			{
				id: '0',
				pageCells: cell('DrawingScale', 2) + cell('PageScale', 1),
				contents: `<Shapes>${shape('1')}</Shapes>`,
			},
			{ id: '1', contents: '<Shapes/>' },
		],
		edit: (zip) => {
			zip.file('custom/opaque.bin', new Uint8Array([0, 1, 255]));
			for (const path of Object.keys(zip.files)) if (zip.files[path]!.dir) delete zip.files[path];
		},
	});

async function withApp(xml: string): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await source());
	const rootRels = await zip.file('_rels/.rels')!.async('string');
	zip.file(
		'_rels/.rels',
		rootRels.replace(
			'</Relationships>',
			'<Relationship Id="app" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',
		),
	);
	zip.file('docProps/app.xml', xml, { createFolders: false });
	return zip.generateAsync({ type: 'uint8array' });
}
const properties = (body: string) =>
	`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${body}</Properties>`;
it('refreshes the Pages property category while preserving other categories and extensions', async () => {
	const app = properties(
		'<HeadingPairs><vt:vector size="6" baseType="variant"><vt:variant><vt:lpstr>Fonts</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Pages</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>Other</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="4" baseType="lpstr"><vt:lpstr>Arial</vt:lpstr><vt:lpstr>Page 1</vt:lpstr><vt:lpstr>Page 2</vt:lpstr><vt:lpstr>opaque</vt:lpstr></vt:vector></TitlesOfParts><Extra xmlns="urn:extra" count="42">keep</Extra>',
	);
	const saved = await editVsdx(await withApp(app), [insert]);
	const zip = await JSZip.loadAsync(saved.bytes),
		xml = await zip.file('docProps/app.xml')!.async('string');
	expect(parseAppProperties(xml)).toMatchObject({
		headingPairs: [
			{ name: 'Fonts', count: 1 },
			{ name: 'Pages', count: 3 },
			{ name: 'Other', count: 1 },
		],
		titlesOfParts: ['Arial', 'Page 1', 'Inserted & Page', 'Page 2', 'opaque'],
	});
	expect(xml).toContain('count="42">keep</Extra>');
	expect(saved.changedParts).toContain('docProps/app.xml');
});
it('refuses inconsistent property vectors rather than erasing unrelated titles', async () => {
	const app = properties(
		'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Pages</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>',
	);
	await expect(editVsdx(await withApp(app), [insert])).rejects.toThrow(/ambiguous category counts/);
});

it('inserts a blank page after its target while preserving scales and untouched payloads', async () => {
	const original = await source(),
		snapshot = new Uint8Array(original);
	const saved = await editVsdx(original, [insert]);
	expect(original).toEqual(snapshot);
	const model = await parseVsdx(saved.bytes);
	expect(model.pages.map((page) => page.id)).toEqual(['0', '99', '1']);
	expect(model.pages[1]).toMatchObject({
		name: 'Inserted & Page',
		width: 4.25,
		height: 5.5,
		drawingToPageScale: 0.5,
		shapes: [],
		isBackground: false,
	});
	const before = await JSZip.loadAsync(original),
		after = await JSZip.loadAsync(saved.bytes);
	for (const path of Object.keys(before.files).filter(
		(path) => !before.files[path]!.dir && !saved.changedParts.includes(path),
	))
		expect(await after.file(path)!.async('uint8array')).toEqual(
			await before.file(path)!.async('uint8array'),
		);
	expect(saved.changedParts).toHaveLength(4);
});
it('supports ordered insertions after a page created in the same atomic transaction', async () => {
	const saved = await editVsdx(await source(), [
		insert,
		{ ...insert, pageId: '100', afterPageId: '99', name: 'Second' },
	]);
	expect((await parseVsdx(saved.bytes)).pages.map((page) => page.id)).toEqual([
		'0',
		'99',
		'100',
		'1',
	]);
});
it.each([
	{ ...insert, pageId: '1' },
	{ ...insert, name: 'page 1' },
	{ ...insert, afterPageId: 'missing' },
	{ ...insert, name: '' },
	{ ...insert, name: 'a\nb' },
	{ ...insert, pageId: '01' },
])('rejects invalid or ambiguous page insertion %j', async (command) => {
	const original = await source(),
		snapshot = new Uint8Array(original);
	await expect(editVsdx(original, [command])).rejects.toThrow();
	expect(original).toEqual(snapshot);
});
it('enforces output and package entry limits and refuses mixed transactions explicitly', async () => {
	const original = await source();
	await expect(editVsdx(original, [insert], { maxOutputBytes: 100 })).rejects.toThrow(
		/output limit/,
	);
	const zip = await JSZip.loadAsync(original);
	await expect(
		editVsdx(original, [insert], {
			limits: { maxEntries: Object.values(zip.files).filter((part) => !part.dir).length },
		}),
	).rejects.toThrow(/entry limit/);
	await expect(
		editVsdx(original, [insert, { type: 'delete-shape', pageId: '0', shapeId: '1' }]),
	).rejects.toThrow(/separate transactions/);
});

const native = process.env.VISIO_NATIVE_PAGE_SCALES_DIR;
describe.skipIf(!native)('native page insertion', () => {
	it('adds a blank scaled page to the native document and permits subsequent shape edits', async () => {
		const original = await readFile(resolve(native!, 'page-scales.vsdx'));
		const saved = await editVsdx(original, [{ ...insert, afterPageId: '6' }]);
		const output = process.env.VISIO_NATIVE_PAGE_INSERT_OUTPUT_DIR;
		if (output) {
			await mkdir(output, { recursive: true });
			await writeFile(resolve(output, 'core-inserted-page.vsdx'), saved.bytes);
		}
		const model = await parseVsdx(saved.bytes);
		const zip = await JSZip.loadAsync(saved.bytes);
		expect(parseAppProperties(await zip.file('docProps/app.xml')!.async('string'))).toMatchObject({
			headingPairs: [{ name: 'Pages', count: 8 }],
			titlesOfParts: model.pages.map((page) => page.name),
		});
		const page = model.pages.find((page) => page.id === '99')!;
		expect(model.pages.indexOf(page) + 1).toBe(nativeEvidence.Index);
		expect(model.pages.length).toBe(nativeEvidence.PageCount);
		expect(page.name).toBe(nativeEvidence.Name);
		expect(page.width).toBeCloseTo(
			(nativeEvidence.PageWidth * nativeEvidence.PageScale) / nativeEvidence.DrawingScale,
			12,
		);
		expect(page.height).toBeCloseTo(
			(nativeEvidence.PageHeight * nativeEvidence.PageScale) / nativeEvidence.DrawingScale,
			12,
		);
		expect(page).toMatchObject({ width: 4.25, height: 5.5, drawingToPageScale: 0.5, shapes: [] });
		const drawn = await editVsdx(saved.bytes, [
			{ type: 'create-rectangle', pageId: '99', shapeId: '1', x: 2, y: 2, width: 2, height: 1 },
		]);
		expect(
			(await parseVsdx(drawn.bytes)).pages.find((page) => page.id === '99')!.shapes[0]!.width,
		).toBe(1);
	});
});
