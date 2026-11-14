import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index';
import { adaptDocumentModel } from './layout/adapter';
import type { LayoutParagraph } from './layout/input';
import { resolveParagraphFormatting } from './paragraph-styles';
import { at } from './test-support/access';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SETTINGS = `<w:settings xmlns:w="${W}"><w:strictFirstAndLastChars/><w:noLineBreaksAfter w:lang="ja-JP" w:val="$"/><w:noLineBreaksBefore w:lang="ja-JP" w:val="!%)"/></w:settings>`;
const STYLES = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:kinsoku w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Chinese"><w:name w:val="Chinese"/><w:rPr><w:lang w:eastAsia="zh-TW"/></w:rPr></w:style></w:styles>`;
const BODY =
	'<w:p><w:pPr><w:kinsoku/></w:pPr><w:r><w:t>一。</w:t></w:r></w:p>' +
	'<w:p><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="false"/></w:pPr><w:r><w:t>二。</w:t></w:r></w:p>' +
	'<w:p><w:pPr><w:pStyle w:val="Chinese"/><w:kinsoku w:val="1"/><w:overflowPunct w:val="0"/></w:pPr><w:r><w:t>三。</w:t></w:r></w:p>';

async function packageBytes(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`,
	);
	zip.file('word/styles.xml', STYLES);
	zip.file('word/settings.xml', SETTINGS);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('kinsoku settings', () => {
	it('reads w:kinsoku through the style hierarchy and the document-wide sets', async () => {
		const { model } = await loadDocx(await packageBytes());
		const paragraphs = model.blocks as Paragraph[];
		const catalog = model.paragraphStyles!;
		expect(
			paragraphs.map((paragraph) => resolveParagraphFormatting(paragraph, catalog).kinsoku),
		).toEqual([true, false, true]);
		expect(model.kinsoku).toEqual({
			strict: true,
			noLineBreaksBefore: { language: 'ja-JP', chars: '!%)' },
			noLineBreaksAfter: { language: 'ja-JP', chars: '$' },
		});
	});

	it('hands the layout each run its East Asian language and the paragraph its switches', async () => {
		const { model } = await loadDocx(await packageBytes());
		const blocks = at(adaptDocumentModel(model).sections, 0).blocks as LayoutParagraph[];
		expect(blocks.map((block) => block.kinsoku)).toEqual([undefined, false, undefined]);
		expect(blocks.map((block) => block.overflowPunct)).toEqual([undefined, undefined, false]);
		expect(blocks.map((block) => [block.autoSpaceDE, block.autoSpaceDN])).toEqual([
			[undefined, undefined],
			[false, false],
			[undefined, undefined],
		]);
		expect(blocks.map((block) => at(block.runs, 0).kinsokuLanguage)).toEqual([
			'ja',
			'ja',
			'zh-Hant',
		]);
		expect(at(blocks, 0).kinsokuSettings).toEqual(model.kinsoku);
	});

	it('saves settings.xml and the paragraph properties unchanged after an edit', async () => {
		const bytes = await packageBytes();
		const loaded = await loadDocx(bytes);
		at((loaded.model.blocks[1] as Paragraph).runs, 0).text = '二、';
		const saved = await JSZip.loadAsync(await loaded.save());
		const original = await JSZip.loadAsync(bytes);
		expect(await saved.file('word/settings.xml')!.async('string')).toContain(
			'<w:noLineBreaksBefore w:lang="ja-JP" w:val="!%)"/>',
		);
		const xml = await saved.file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:kinsoku/>');
		expect(xml).toContain('<w:kinsoku w:val="1"/><w:overflowPunct w:val="0"/>');
		expect(xml).toContain('<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="false"/>');
		expect(await saved.file('word/styles.xml')!.async('string')).toBe(
			await original.file('word/styles.xml')!.async('string'),
		);
	});
});
