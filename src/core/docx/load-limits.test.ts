import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { loadDocument } from './load/detect';
import { loadDocx } from './parse';
import { docx, p } from './test-support/compare-docs';
import { parseXml } from './xml';

afterEach(() => vi.restoreAllMocks());

describe('docx load limits', () => {
	it('rejects a main document part over limits.maxDocumentPartBytes and loads it when raised', async () => {
		const bytes = await docx(p('x'.repeat(4000)));
		await expect(loadDocx(bytes, { limits: { maxDocumentPartBytes: 2000 } })).rejects.toThrow(
			'word/document.xml exceeds the 2000 bytes limit (limits.maxDocumentPartBytes)',
		);
		await expect(loadDocument(bytes, { limits: { maxDocumentPartBytes: 2000 } })).rejects.toThrow(
			/maxDocumentPartBytes/,
		);
		expect(
			(await loadDocx(bytes, { limits: { maxDocumentPartBytes: 8000 } })).model.blocks,
		).toHaveLength(1);
	});

	it('refuses a limit that would switch a guard off', async () => {
		const bytes = await docx(p('a'));
		for (const value of [0, -1, Number.NaN, Number.POSITIVE_INFINITY])
			await expect(loadDocx(bytes, { limits: { maxParts: value } })).rejects.toThrow(RangeError);
	});

	it('saves a package loaded under raised limits with the same limits', async () => {
		const zip = await JSZip.loadAsync(await docx(p('kept')));
		for (let index = 0; index < 5001; index++) zip.file(`customXml/item${index}.xml`, '<a/>');
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		await expect(loadDocx(bytes)).rejects.toThrow();
		const loaded = await loadDocx(bytes, { limits: { maxParts: 6000 } });
		const first = loaded.model.blocks[0]!;
		const edited = { ...loaded.model, blocks: [{ ...first, runs: [{ text: 'edited' }] }] };
		const saved = await loaded.save(edited as typeof loaded.model);
		const reloaded = await loadDocx(saved, { limits: { maxParts: 6000 } });
		expect(reloaded.model.blocks[0]).toMatchObject({ runs: [{ text: 'edited' }] });
	});
});

describe('docx load scans', () => {
	it('walks document.xml once for every feature it reports', async () => {
		const body =
			'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>HYPERLINK "x"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:customXml><w:r><w:t>tag</w:t></w:r></w:customXml><w:r><w:br w:type="textWrapping" w:clear="all"/></w:r></w:p>';
		const bytes = await docx(body);
		const proto = Object.getPrototypeOf(parseXml('<a/>')) as Document;
		const walks = vi.spyOn(proto, 'getElementsByTagNameNS');
		const { model } = await loadDocx(bytes);
		const onMainDocument = walks.mock.contexts.filter(
			(context) => (context as Document).documentElement?.localName === 'document',
		);
		expect(onMainDocument).toHaveLength(0);
		expect(model.warnings.join('\n')).toMatch(/Fields show the result Word last saved/);
		expect(model.warnings.join('\n')).toMatch(/Complex HYPERLINK field codes/);
		expect(model.warnings.join('\n')).toMatch(/Smart tags and custom XML markup/);
	});
});
