import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx, saveDocx } from './index';
import { expectParagraph } from './test-support/access';
import { first, getW, parseXml, WORD_NS } from './xml';

async function fixture() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}" xmlns:x="urn:opaque-run"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="eastAsia" w:asciiTheme="minorAscii"/><w:b w:val="0"/><x:opaque x:value="first"/></w:rPr><w:t>First</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" w:hint="default"/><w:color w:val="FF0000" w:themeColor="accent1"/><x:opaque x:value="second"/></w:rPr><w:t>Second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

describe('opaque current run properties', () => {
	it('preserves native Word outline and shadow through tracked text export', async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(
					new URL('./__fixtures__/opaque-run-properties/outline-shadow.docx', import.meta.url),
				),
			),
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const original = paragraph.runs[0]!;
		paragraph.runs = [
			{ ...original, text: original.text.slice(0, 4) },
			{
				...original,
				text: '!',
				revision: {
					id: 'opaque-text',
					kind: 'insert',
					author: 'Codex',
					date: '2026-10-07T00:00:00Z',
				},
			},
			{ ...original, text: original.text.slice(4) },
		];
		for (const bytes of [await loaded.save(loaded.model), await saveDocx(loaded.model)]) {
			const reopened = await loadDocx(bytes);
			const runs = expectParagraph(reopened.model.blocks[0]).runs;
			expect(runs.map((run) => run.text).join('')).toBe('Outl!ined text');
			for (const run of runs) {
				const props = parseXml(run.sourceRunPropertiesXml!).documentElement;
				expect(first(props, 'outline')).toBeDefined();
				expect(first(props, 'shadow')).toBeDefined();
			}
		}
	});
	it('preserves distinct property bases when tracked edits split and reorder runs', async () => {
		const loaded = await fixture();
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const [a, b] = paragraph.runs;
		paragraph.runs = [
			{ ...a!, text: 'Fi' },
			{ ...a!, text: '!', revision: { id: 'edit', kind: 'insert', author: 'Ada' } },
			{ ...b!, text: 'Second' },
			{ ...a!, text: 'rst' },
		];
		for (const bytes of [await loaded.save(loaded.model), await saveDocx(loaded.model)]) {
			const reopened = await loadDocx(bytes);
			const runs = expectParagraph(reopened.model.blocks[0]).runs;
			expect(runs.map((run) => run.text)).toEqual(['Fi', '!', 'Second', 'rst']);
			for (let i = 0; i < runs.length; i++) {
				const props = parseXml(runs[i]!.sourceRunPropertiesXml!).documentElement;
				const fonts = first(props, 'rFonts');
				expect(getW(fonts, 'hint')).toBe(i === 2 ? 'default' : 'eastAsia');
				expect(
					props
						.getElementsByTagNameNS('urn:opaque-run', 'opaque')[0]
						?.getAttributeNS('urn:opaque-run', 'value'),
				).toBe(i === 2 ? 'second' : 'first');
			}
			expect(runs[1]?.revision?.kind).toBe('insert');
		}
	});

	it('keeps safety guards when a caller drops the complete source property basis', async () => {
		const loaded = await fixture();
		const original = await loaded.save();
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs = [{ text: 'Fi' }, { text: 'rst' }, { text: 'Second' }];
		await expect(loaded.save(loaded.model)).rejects.toThrow('changing run boundaries');
		expect(await (await loadDocx(original)).save()).toEqual(original);
	});

	it('retains unmodeled attributes on formatting history when its run splits', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${WORD_NS}" xmlns:x="urn:opaque-run"><w:body><w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Ada" x:owner="retained"><w:rPr/></w:rPrChange></w:rPr><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const original = paragraph.runs[0]!;
		paragraph.runs = [
			{ ...original, text: 'Te' },
			{ ...original, text: 'xt' },
		];
		const xml = await (
			await JSZip.loadAsync(await loaded.save(loaded.model))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml.match(/x:owner="retained"/g)).toHaveLength(2);
	});
});
