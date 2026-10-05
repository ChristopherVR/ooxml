import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { editVsdx } from './edit.js';
import { fixture, shape, relations } from './test-fixtures.js';
const command = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'New' };
const input = (edit?: (zip: JSZip) => void) =>
	fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', '<Text>Old</Text>')}</Shapes>` }],
		...(edit ? { edit } : {}),
	});
describe('VSDX edit security boundaries', () => {
	it.each(['\u0000', '\u0001', '\uffff', '\ud800', '\udc00', 'a\rb'])(
		'rejects invalid/unsupported replacement characters',
		async (text) => {
			await expect(editVsdx(await input(), [{ ...command, text }])).rejects.toMatchObject({
				code: 'INVALID_EDIT_TEXT',
			});
		},
	);
	it('enforces command, replacement, output, entry and total limits', async () => {
		const bytes = await input();
		await expect(editVsdx(bytes, [command, command], { maxEdits: 1 })).rejects.toMatchObject({
			code: 'LIMIT_EDITS',
		});
		await expect(editVsdx(bytes, [command], { maxTextCharacters: 2 })).rejects.toMatchObject({
			code: 'LIMIT_EDIT_TEXT',
		});
		await expect(editVsdx(bytes, [command], { maxOutputBytes: 10 })).rejects.toMatchObject({
			code: 'LIMIT_EDIT_OUTPUT',
		});
		await expect(editVsdx(bytes, [], { maxOutputBytes: 10 })).rejects.toMatchObject({
			code: 'LIMIT_EDIT_OUTPUT',
		});
		await expect(
			editVsdx(bytes, [{ ...command, text: 'a'.repeat(2000) }], {
				limits: { maxEntryBytes: 1000 },
			}),
		).rejects.toMatchObject({ code: 'LIMIT_ENTRY' });
		await expect(
			editVsdx(bytes, [{ ...command, text: 'a'.repeat(10000) }], {
				limits: { maxTotalBytes: 5000 },
			}),
		).rejects.toMatchObject({ code: 'LIMIT_TOTAL' });
	});
	it('enforces per-operation deadline', async () => {
		const bytes = await input();
		const now = Date.now(),
			spy = vi
				.spyOn(Date, 'now')
				.mockImplementationOnce(() => now)
				.mockReturnValue(now + 100);
		try {
			await expect(
				editVsdx(bytes, [command], { limits: { maxRuntimeMs: 10 } }),
			).rejects.toMatchObject({ code: 'LIMIT_RUNTIME' });
		} finally {
			spy.mockRestore();
		}
	});
	it('rejects corrupt unused opaque payload CRC', async () => {
		const bytes = await input((zip) =>
			zip.file('unused.bin', new Uint8Array([41, 42, 43, 44, 45, 46]), { compression: 'STORE' }),
		);
		const needle = new Uint8Array([41, 42, 43, 44, 45, 46]);
		const at = bytes.findIndex((_, i) => needle.every((value, j) => bytes[i + j] === value));
		expect(at).toBeGreaterThan(0);
		bytes[at] = 0;
		await expect(editVsdx(bytes, [command])).rejects.toMatchObject({ code: 'ZIP_MISMATCH' });
	});
	it('rejects unsafe names, excessive compression and XML entities', async () => {
		await expect(
			editVsdx(await input((zip) => zip.file('../bad', 'x')), [command]),
		).rejects.toMatchObject({ code: 'UNSAFE_PATH' });
		const zip = await JSZip.loadAsync(await input((zip) => zip.file('bomb', 'a'.repeat(100000))));
		await expect(
			editVsdx(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), [command]),
		).rejects.toMatchObject({ code: 'LIMIT_RATIO' });
		await expect(
			editVsdx(
				await input((zip) =>
					zip.file(
						'visio/pages/page1.xml',
						'<!DOCTYPE PageContents [<!ENTITY x "bad">]><PageContents/>',
					),
				),
				[command],
			),
		).rejects.toMatchObject({ code: 'INVALID_XML' });
	});
	it('rejects signed packages by path, content type and relationship', async () => {
		for (const change of [
			(zip: JSZip) => {
				zip.file('_xmlsignatures/sig.xml', '<Signature/>');
			},
			(zip: JSZip) => {
				zip.file('signature.bin', 'opaque');
				zip.file(
					'visio/_rels/document.xml.rels',
					relations(
						'<Relationship Id="sig" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="../signature.bin"/>',
					),
				);
			},
			(zip: JSZip) => {
				zip.file(
					'[Content_Types].xml',
					'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/><Default Extension="sig" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/></Types>',
				);
			},
		])
			await expect(editVsdx(await input(change), [command])).rejects.toMatchObject({
				code: 'UNSUPPORTED_EDIT_PACKAGE',
			});
	});
	it.each([
		'<Shapes><Shape ID="1"><Text>Old</Text></Shape><Shape ID="1"><Text>Other</Text></Shape></Shapes>',
		'<Shapes><Shape ID="1"><Text>Old</Text><Unknown>before&#13;after</Unknown></Shape></Shapes>',
		'<Shapes><Shape ID="1" Extra="before&#10;after"><Text>Old</Text></Shape></Shapes>',
	])('fails closed on ambiguous IDs or serializer-normalized content', async (contents) => {
		await expect(
			editVsdx(await fixture({ pages: [{ id: '0', contents }] }), [command]),
		).rejects.toMatchObject({
			code: contents.includes('Other') ? 'INVALID_SHAPE_ID' : 'UNSUPPORTED_XML_EDIT',
		});
	});
});
