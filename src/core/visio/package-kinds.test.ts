import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { readVisioStencil } from './stencil-file';
import { fixture, relation, relations, shape } from './test-fixtures';

const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MAIN = 'application/vnd.ms-visio.';
const VBA = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 250, 251, 252]);
const command = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'New' };

async function kind(
	mainType: string,
	options: { macros?: boolean; signed?: boolean } = {},
): Promise<Uint8Array> {
	return fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', '<Text>Old</Text>')}</Shapes>` }],
		masters: [{ id: '1', shapes: shape('1') }],
		edit(zip) {
			const macro = options.macros
				? `<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>`
				: '';
			const signed = options.signed
				? `<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/><Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>`
				: '';
			zip.file(
				'[Content_Types].xml',
				`<Types xmlns="${TYPES}">${macro}${signed}<Override PartName="/visio/document.xml" ContentType="${mainType}"/></Types>`,
			);
			if (options.macros) {
				zip.file('visio/vbaProject.bin', VBA);
				zip.file(
					'visio/_rels/document.xml.rels',
					relations(
						relation('rId1', 'pages', 'pages/pages.xml') +
							relation('rId2', 'masters', 'masters/masters.xml') +
							'<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>',
					),
				);
			}
			if (options.signed) {
				zip.file('_xmlsignatures/origin.sigs', '');
				zip.file('_xmlsignatures/sig1.xml', '<Signature/>');
				zip.file(
					'_xmlsignatures/_rels/origin.sigs.rels',
					relations(
						'<Relationship Id="s1" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>',
					),
				);
				zip.file(
					'_rels/.rels',
					relations(
						relation('rId1', 'document', 'visio/document.xml') +
							'<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>',
					),
				);
			}
		},
	});
}
const read = async (bytes: Uint8Array, path: string) =>
	JSZip.loadAsync(bytes).then((zip) => zip.file(path)?.async('string'));

describe('Visio package kinds', () => {
	it.each([
		['.vsdx', `${MAIN}drawing.main+xml`],
		['.vsdm', `${MAIN}drawing.macroEnabled.main+xml`],
		['.vstx', `${MAIN}template.main+xml`],
		['.vstm', `${MAIN}template.macroEnabled.main+xml`],
	])('opens %s as a drawing', async (_name, type) => {
		const document = await parseVsdx(await kind(type));
		expect(document.pages).toHaveLength(1);
		expect(document.pages[0]!.shapes).toHaveLength(1);
	});

	it.each([
		['.vssx', `${MAIN}stencil.main+xml`],
		['.vssm', `${MAIN}stencil.macroEnabled.main+xml`],
	])('opens %s for its masters but not as a drawing', async (_name, type) => {
		const bytes = await kind(type);
		expect((await readVisioStencil(bytes)).masters).toHaveLength(1);
		await expect(parseVsdx(bytes)).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' });
	});

	it('reports that a VBA project is present and never runs it', async () => {
		const document = await parseVsdx(
			await kind(`${MAIN}drawing.macroEnabled.main+xml`, { macros: true }),
		);
		expect(document.diagnostics.find((item) => item.code === 'macros-not-run')?.severity).toBe(
			'info',
		);
	});

	it('keeps the VBA project, its relationship and the macro-enabled type through an edit', async () => {
		const type = `${MAIN}drawing.macroEnabled.main+xml`;
		const source = await kind(type, { macros: true });
		const result = await editVsdx(source, [command]);
		const zip = await JSZip.loadAsync(result.bytes);
		expect([...(await zip.file('visio/vbaProject.bin')!.async('uint8array'))]).toEqual([...VBA]);
		expect(await read(result.bytes, 'visio/_rels/document.xml.rels')).toContain('vbaProject.bin');
		expect(await read(result.bytes, '[Content_Types].xml')).toBe(
			await read(source, '[Content_Types].xml'),
		);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.text.plainText).toContain('New');
	});

	it('saves a template as a drawing of the same kind, keeping macros', async () => {
		const plain = await editVsdx(await kind(`${MAIN}template.main+xml`), [command]);
		expect(await read(plain.bytes, '[Content_Types].xml')).toContain(`${MAIN}drawing.main+xml`);
		expect(plain.diagnostics.map((item) => item.code)).toContain('TEMPLATE_SAVED_AS_DRAWING');
		const macro = await editVsdx(
			await kind(`${MAIN}template.macroEnabled.main+xml`, { macros: true }),
			[command],
		);
		expect(await read(macro.bytes, '[Content_Types].xml')).toContain(
			`${MAIN}drawing.macroEnabled.main+xml`,
		);
		const zip = await JSZip.loadAsync(macro.bytes);
		expect(zip.file('visio/vbaProject.bin')).toBeTruthy();
	});

	it('drops digital signatures on an edit and says so', async () => {
		const source = await kind(`${MAIN}drawing.main+xml`, { signed: true });
		const result = await editVsdx(source, [command]);
		const zip = await JSZip.loadAsync(result.bytes);
		expect(Object.keys(zip.files).some((name) => name.startsWith('_xmlsignatures'))).toBe(false);
		expect(await read(result.bytes, '_rels/.rels')).not.toContain('digital-signature');
		expect(await read(result.bytes, '[Content_Types].xml')).not.toMatch(/digital-signature|sigs/);
		expect(result.diagnostics.map((item) => item.code)).toContain('SIGNATURE_REMOVED');
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.text.plainText).toContain('New');
	});

	it('drops signatures of a signed macro-enabled package and keeps its macros', async () => {
		const source = await kind(`${MAIN}drawing.macroEnabled.main+xml`, {
			macros: true,
			signed: true,
		});
		const result = await editVsdx(source, [command]);
		const zip = await JSZip.loadAsync(result.bytes);
		expect(zip.file('visio/vbaProject.bin')).toBeTruthy();
		expect(zip.file('_xmlsignatures/sig1.xml')).toBeNull();
	});
});
