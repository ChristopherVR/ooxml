import { it, expect } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit.js';
import { fixture, shape, xml } from './test-fixtures.js';
const command = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'new' };
const base = (
	contents = `<Shapes>${shape('1', '<Text>old</Text>')}</Shapes>`,
	edit?: (zip: JSZip) => void,
) => fixture({ pages: [{ id: '0', contents }], ...(edit ? { edit } : {}) });
it('review fails closed rather than changing unrelated CR text', async () => {
	const input = await base(
		`<Shapes>${shape('1', '<Text>old</Text>')}${shape('2', '<Text>a&#13;b</Text>')}</Shapes>`,
	);
	const original = input.slice();
	await expect(editVsdx(input, [command])).rejects.toMatchObject({ code: 'UNSUPPORTED_XML_EDIT' });
	expect(input).toEqual(original);
});
it('review rejects duplicate Shapes containers', async () => {
	const input = await base(
		`<Shapes>${shape('1', '<Text>old</Text>')}</Shapes><Shapes>${shape('1', '<Text>other</Text>')}</Shapes>`,
	);
	await expect(editVsdx(input, [command])).rejects.toBeDefined();
});
it('review rejects spoofed relationship prefix', async () => {
	const input = await base(undefined, (z) =>
		z.file(
			'visio/pages/pages.xml',
			xml('Pages', '<Page ID="0"><Rel xmlns:r="urn:evil" r:id="rId1"/></Page>'),
		),
	);
	await expect(editVsdx(input, [command])).rejects.toBeDefined();
});
it('review rejects percent encoded signature path', async () => {
	const input = await base(undefined, (z) => z.file('%5Fxmlsignatures/sig.xml', '<Signature/>'));
	await expect(editVsdx(input, [command])).rejects.toBeDefined();
});
it('review rejects unadvertised VBA payload', async () => {
	const input = await base(undefined, (z) =>
		z.file('visio/vbaProject.bin', new Uint8Array([1, 2, 3])),
	);
	await expect(editVsdx(input, [command])).rejects.toBeDefined();
});
it('review rejects spoofed ContentType override namespace', async () => {
	const input = await base(undefined, (z) =>
		z.file(
			'[Content_Types].xml',
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><x:Override xmlns:x="urn:evil" PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/></Types>',
		),
	);
	expect(
		await editVsdx(input, [command]).then(
			() => false,
			() => true,
		),
	).toBe(true);
});
it('review rejects duplicated Page Rel', async () => {
	const input = await base(undefined, (z) =>
		z.file(
			'visio/pages/pages.xml',
			xml('Pages', '<Page ID="0"><Rel r:id="rId1"/><Rel r:id="rId1"/></Page>'),
		),
	);
	expect(
		await editVsdx(input, [command]).then(
			() => false,
			() => true,
		),
	).toBe(true);
});
it('review enforces final aggregate XML budget', async () => {
	const { inspectXml } = await import('./xml-validation.js');
	const { DEFAULTS } = await import('./package-common.js');
	const input = await base(`<Shapes>${shape('1', '<Text/>')}</Shapes>`),
		zip = await JSZip.loadAsync(input);
	let total = 0;
	for (const file of Object.values(zip.files))
		if (!file.dir) total += inspectXml(await file.async('string'), DEFAULTS, () => {});
	expect(
		await editVsdx(input, [command], { limits: { maxTotalXmlNodes: total } }).then(
			() => false,
			(e) => e.code,
		),
	).toBe('LIMIT_XML_TOTAL');
});
