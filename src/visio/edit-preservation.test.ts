import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit.js';
import { VisioPackage } from './package.js';
import { fixture, shape, ns } from './test-fixtures.js';
const edit = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'Changed' };
describe('VSDX edit preservation', () => {
	it('preserves every untouched part byte-for-byte and unknown edited-page content semantically', async () => {
		const input = await fixture({
			pages: [
				{
					id: '0',
					contents: `<!--page note--><?tool keep?><Shapes>${shape('1', '<Cell N="Width" V="3" F="1+2"/><Text>Old</Text><x:Extension xmlns:x="urn:example" x:flag="yes">&amp;unknown</x:Extension>', 'Custom="kept"')}</Shapes>`,
				},
			],
			edit(zip) {
				zip.file('custom/opaque.bin', new Uint8Array([0, 255, 38, 12]));
				zip.file(
					'custom/opaque.xml',
					'<?xml version="1.0"?><alien weird=" untouched "><![CDATA[content]]></alien>',
				);
				zip.file(
					'custom/_rels/opaque.xml.rels',
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="ext" Type="urn:external" Target="https://example.invalid/never-fetch" TargetMode="External"/></Relationships>',
				);
			},
		});
		const output = await editVsdx(input, [edit]);
		const before = await VisioPackage.open(input),
			after = await VisioPackage.open(output.bytes);
		expect(after.paths().sort()).toEqual(before.paths().sort());
		for (const path of before.paths())
			if (!output.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		const root = await after.readXml('visio/pages/page1.xml');
		expect(root.getElementsByTagNameNS('urn:example', 'Extension')[0]?.textContent).toBe(
			'&unknown',
		);
		expect(root.getElementsByTagNameNS(ns, 'Shape')[0]?.getAttribute('Custom')).toBe('kept');
		expect(root.getElementsByTagNameNS(ns, 'Cell')[0]?.getAttribute('F')).toBe('1+2');
		const xml = new TextDecoder().decode(await after.readBytes('visio/pages/page1.xml'));
		expect(xml).toContain('<!--page note-->');
		expect(xml).toContain('<?tool keep?>');
	});
	it('resolves noncanonical page paths through relationships', async () => {
		const zip = await JSZip.loadAsync(
			await fixture({
				pages: [{ id: '0', contents: `<Shapes>${shape('1', '<Text>Old</Text>')}</Shapes>` }],
			}),
		);
		zip.file('custom/page.xml', await zip.file('visio/pages/page1.xml')!.async('uint8array'));
		zip.remove('visio/pages/page1.xml');
		const rel = 'visio/pages/_rels/pages.xml.rels';
		zip.file(
			rel,
			(await zip.file(rel)!.async('string')).replace('page1.xml', '../../custom/page.xml'),
		);
		expect(
			(await editVsdx(await zip.generateAsync({ type: 'uint8array' }), [edit])).changedParts,
		).toEqual(['custom/page.xml']);
	});
	it.each(['utf-16le', 'utf-16be'])(
		'converts edited %s XML to declared UTF-8',
		async (encoding) => {
			const input = await fixture({
				edit(zip) {
					const xml = `<?xml version="1.0" encoding="UTF-16"?><v:PageContents xmlns:v="${ns}"><v:Shapes><v:Shape ID="1"><v:Text>Old</v:Text></v:Shape></v:Shapes></v:PageContents>`;
					const bytes = new Uint8Array(2 + xml.length * 2),
						view = new DataView(bytes.buffer);
					view.setUint16(0, 0xfeff, encoding === 'utf-16le');
					for (let i = 0; i < xml.length; i++)
						view.setUint16(2 + 2 * i, xml.charCodeAt(i), encoding === 'utf-16le');
					zip.file('visio/pages/page1.xml', bytes);
				},
			});
			const result = await editVsdx(input, [edit]);
			const pkg = await VisioPackage.open(result.bytes);
			expect((await pkg.readXml('visio/pages/page1.xml')).textContent).toBe('Changed');
			expect(new TextDecoder().decode(await pkg.readBytes('visio/pages/page1.xml'))).toContain(
				'encoding="UTF-8"',
			);
		},
	);
});
