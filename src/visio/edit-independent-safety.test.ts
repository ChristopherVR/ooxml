import { it, expect } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit.js';
import { fixture, shape } from './test-fixtures.js';
import { VisioPackage } from './package.js';
const command = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'new' };
const base = () =>
	fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', '<Text>old</Text>')}</Shapes>` }],
		edit: (z) => {
			z.file('custom/opaque.bin', new Uint8Array([0, 255, 1, 128, 42]));
			z.file(
				'custom/unknown.xml',
				'\ufeff<?xml version="1.0"?><unknown a="  x  "><![CDATA[a<b]]></unknown>',
			);
		},
	});
it('independent exact part payload preservation', async () => {
	const a = await base(),
		b = (await editVsdx(a, [command])).bytes;
	const az = await JSZip.loadAsync(a),
		bz = await JSZip.loadAsync(b);
	for (const [name, file] of Object.entries(az.files))
		if (!file.dir && name !== 'visio/pages/page1.xml')
			expect(await bz.file(name)!.async('uint8array')).toEqual(await file.async('uint8array'));
});
it('independent shared buffer is copied before await', async () => {
	const a = await base(),
		buf = new Uint8Array(new SharedArrayBuffer(a.length));
	buf.set(a);
	const pending = editVsdx(buf, [command]);
	buf.fill(0);
	const pkg = await VisioPackage.open((await pending).bytes);
	expect((await pkg.readXml('visio/pages/page1.xml')).textContent).toBe('new');
});
it('independent resizable buffer is copied before await', async () => {
	const a = await base();
	const buf = new (ArrayBuffer as unknown as {
		new (
			length: number,
			options: {
				maxByteLength: number;
			},
		): ArrayBuffer & {
			resize(length: number): void;
		};
	})(a.length, { maxByteLength: a.length * 2 });
	new Uint8Array(buf).set(a);
	const pending = editVsdx(buf, [command]);
	buf.resize(0);
	const pkg = await VisioPackage.open((await pending).bytes);
	expect((await pkg.readXml('visio/pages/page1.xml')).textContent).toBe('new');
});
it('independent invalid text is rejected atomically', async () => {
	const a = await base(),
		copy = a.slice();
	for (const text of ['\u0000', '\ud800', '\udfff', '\ufffe'])
		expect(
			await editVsdx(a, [{ ...command, text }]).then(
				() => false,
				() => true,
			),
		).toBe(true);
	expect(a).toEqual(copy);
});
it('independent caps output', async () => {
	expect(
		await editVsdx(await base(), [command], { maxOutputBytes: 10 }).then(
			() => false,
			(e) => e.code,
		),
	).toBe('LIMIT_EDIT_OUTPUT');
});
