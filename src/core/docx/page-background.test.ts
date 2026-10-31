import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { saveDocx } from './save';
import { createDocument } from './model';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const CONTENT_TYPES =
	'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';

const themeXml =
	'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Native"><a:accent2><a:srgbClr val="ED7D31"/></a:accent2></a:clrScheme></a:themeElements></a:theme>';
async function packageWith(background: string, themed = false): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', CONTENT_TYPES);
	if (themed) zip.file('word/theme/theme1.xml', themeXml);
	zip.file(
		'word/document.xml',
		`<w:document ${NS}>${background}<w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;

describe('page colour', () => {
	it('replaces a theme fill when the selected literal equals its stale cached RGB', async () => {
		const loaded = await loadDocx(
			await packageWith('<w:background w:color="FF0000" w:themeColor="accent2"/>', true),
		);
		const saved = await loaded.save({ ...loaded.model, pageColor: 'FF0000' });
		expect(await documentXml(saved)).toContain('<w:background w:color="FF0000"/>');
		expect(await documentXml(saved)).not.toContain('w:themeColor');
		expect((await loadDocx(saved)).model.pageColor).toBe('FF0000');
	});
	it('resolves theme-only fills and ignores a stale literal cache, retaining the source on edits', async () => {
		for (const cache of ['', 'w:color="FF0000"']) {
			const source = `<w:background ${cache} w:themeColor="accent2" w:themeTint="7F" w:themeShade="40"/>`;
			const loaded = await loadDocx(await packageWith(source, true));
			expect(loaded.model.pageColor).toBe('F6BE98');
			const blocks = loaded.model.blocks.map((block) =>
				block.type === 'paragraph' ? { ...block, runs: [{ text: 'Edited text' }] } : block,
			);
			const saved = await loaded.save({ ...loaded.model, blocks });
			expect(await documentXml(saved)).toContain('w:themeTint="7F"');
			expect(await documentXml(saved)).toContain('w:themeShade="40"');
			if (cache) expect(await documentXml(saved)).toContain('w:color="FF0000"');
			expect((await loadDocx(saved)).model.pageColor).toBe('F6BE98');
		}
	});

	it('uses the logical background colour mapping and retains a fallback if the theme is incomplete', async () => {
		const zip = await JSZip.loadAsync(
			await packageWith('<w:background w:themeColor="background1"/>', true),
		);
		zip.file(
			'word/settings.xml',
			`<w:settings ${NS}><w:clrSchemeMapping w:bg1="accent2"/></w:settings>`,
		);
		expect((await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model.pageColor).toBe(
			'ED7D31',
		);
		expect(
			(
				await loadDocx(
					await packageWith('<w:background w:color="00FF00" w:themeColor="accent6"/>', true),
				)
			).model.pageColor,
		).toBe('00FF00');
	});
	it('reads w:background, ignoring auto and unset fills', async () => {
		expect(
			(await loadDocx(await packageWith('<w:background w:color="ffcc00"/>'))).model.pageColor,
		).toBe('FFCC00');
		expect(
			(await loadDocx(await packageWith('<w:background w:color="auto"/>'))).model.pageColor,
		).toBeUndefined();
		expect((await loadDocx(await packageWith(''))).model.pageColor).toBeUndefined();
	});

	it('writes a new colour before the body and asks settings.xml to display it', async () => {
		const loaded = await loadDocx(await packageWith(''));
		const saved = await loaded.save({ ...loaded.model, pageColor: '#e2efd9' });
		const xml = await documentXml(saved);
		expect(xml).toMatch(/<w:background w:color="E2EFD9"\/><w:body>/);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('word/settings.xml')!.async('string')).toContain(
			'<w:displayBackgroundShape/>',
		);
		expect((await loadDocx(saved)).model.pageColor).toBe('E2EFD9');
	});

	it('leaves an unchanged colour byte-identical and removes a cleared one', async () => {
		const source = '<w:background w:color="FFCC00" w:themeColor="accent4" w:themeTint="99"/>';
		const loaded = await loadDocx(await packageWith(source));
		expect(await documentXml(await loaded.save(loaded.model))).toContain(source);
		const cleared = { ...loaded.model };
		delete cleared.pageColor;
		expect(await documentXml(await loaded.save(cleared))).not.toContain('w:background');
	});

	it('saves a new document with a colour', async () => {
		const model = { ...createDocument(), pageColor: '00B0F0' };
		const xml = await documentXml(await saveDocx(model));
		expect(xml).toContain('<w:background w:color="00B0F0"/>');
	});

	it('rejects a colour that is not a hex value', async () => {
		const loaded = await loadDocx(await packageWith(''));
		await expect(loaded.save({ ...loaded.model, pageColor: 'blue' })).rejects.toThrow();
	});
});
