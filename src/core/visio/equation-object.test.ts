import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { parseVsdx } from './parser';

const fixture = () =>
	readFile(new URL('./__fixtures__/word-equation-native.vsdx', import.meta.url));

it('reads equations in a desktop Visio Word OLE object and preserves the package and cached preview on edit', async () => {
	const native = await fixture();
	const model = await parseVsdx(native);
	const page = model.pages[0]!;
	const shape = page.shapes[0]!;
	expect(shape.kind).toBe('foreign');
	expect(shape.image?.mimeType).toBe('image/svg+xml');
	const svg = new TextDecoder().decode(shape.image?.bytes);
	expect(svg).toContain('<line');
	expect(svg).toContain('Cambria Math');
	expect(model.diagnostics.map((d) => d.code)).not.toContain('unsupported-foreign-object');
	const edited = await editVsdx(native, [
		{ type: 'move-shape', pageId: page.id, shapeId: shape.id, x: 4, y: 5 },
	]);
	const before = await VisioPackage.open(native);
	const after = await VisioPackage.open(edited.bytes);
	for (const part of [
		'visio/embeddings/Microsoft_Word_Document.docx',
		'visio/embeddings/_rels/Microsoft_Word_Document.docx.rels',
		'visio/media/image1.emf',
	])
		expect(await after.readBytes(part)).toEqual(await before.readBytes(part));
	const reopened = (await parseVsdx(edited.bytes)).pages[0]?.shapes[0];
	expect(reopened?.image?.bytes).toEqual(shape.image?.bytes);
	expect(reopened?.transform).not.toEqual(shape.transform);
});
