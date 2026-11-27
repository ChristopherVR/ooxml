import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx } from './index';
import { createDocument, type DocumentModel, type InlineImage, type Paragraph } from './model';
import { newShape } from './shape-model';
import { expectParagraph } from './test-support/access';
import { blocksArePlainLines } from './write-text-box-blocks';

const withImage = (image: InlineImage): DocumentModel => {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: '', image }] }];
	return model;
};
const xmlOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;
const boxBlocks: Paragraph[] = [
	{ type: 'paragraph', id: 'b0', runs: [{ text: 'Bold', bold: true }, { text: ' lead' }] },
	{ type: 'paragraph', id: 'b1', align: 'right', runs: [{ text: 'Second' }] },
];

describe('blocks of a new text box', () => {
	it('writes the blocks and body settings of a new shape instead of its plain lines', async () => {
		const bytes = await saveDocx(
			withImage({
				relId: '',
				partName: '',
				contentType: 'application/octet-stream',
				widthPx: 192,
				heightPx: 96,
				unsupported: 'Shape',
				shape: { ...newShape('rect'), text: ['Bold lead', 'Second'] },
				textBoxContent: {
					blocks: boxBlocks,
					anchor: 'b',
					insetLeftEmu: 0,
					insetTopEmu: 36000,
					insetRightEmu: 0,
					insetBottomEmu: 36000,
					autofit: 'shape',
				},
			}),
		);
		const xml = await xmlOf(bytes);
		const story = xml.slice(xml.indexOf('<w:txbxContent'), xml.indexOf('</w:txbxContent>'));
		expect(story).toContain('<w:b/>');
		expect(story).toContain('<w:jc w:val="right"/>');
		expect(story).not.toContain('<w:jc w:val="center"/>');
		expect(xml).toMatch(/<wps:bodyPr [^>]*lIns="0" tIns="36000" rIns="0" bIns="36000" anchor="b"/);
		expect(xml).toContain('<a:spAutoFit/>');
		expect(xml).not.toContain('<a:noAutofit/>');
		const reopened = await loadDocx(bytes);
		const image = expectParagraph(reopened.model.blocks[0]).runs[0]?.image;
		const blocks = image?.textBoxContent?.blocks as Paragraph[] | undefined;
		expect(blocks?.map((block) => block.runs.map((run) => run.text).join(''))).toEqual([
			'Bold lead',
			'Second',
		]);
		expect(blocks?.[0]?.runs[0]).toMatchObject({ text: 'Bold', bold: true });
		expect(image?.textBoxContent).toMatchObject({ anchor: 'b', autofit: 'shape' });
	});

	it('writes rich blocks into a new simple text box, and keeps plain lines in the simple form', async () => {
		const box = (blocks: Paragraph[], lines: string[]): InlineImage => ({
			relId: '',
			partName: '',
			contentType: 'application/octet-stream',
			widthPx: 192,
			heightPx: 96,
			textBoxEditable: true,
			textBoxText: lines,
			textBoxContent: { blocks },
		});
		const rich = await xmlOf(await saveDocx(withImage(box(boxBlocks, ['Bold lead', 'Second']))));
		expect(rich).toContain('<w:b/>');
		const plainBlocks: Paragraph[] = [{ type: 'paragraph', id: 'b0', runs: [{ text: 'Only' }] }];
		expect(blocksArePlainLines(box(plainBlocks, ['Only']))).toBe(true);
		expect(blocksArePlainLines(box(boxBlocks, ['Bold lead', 'Second']))).toBe(false);
		const plain = await xmlOf(await saveDocx(withImage(box(plainBlocks, ['Only']))));
		expect(plain).toContain('<w:txbxContent><w:p><w:r><w:t>Only</w:t></w:r></w:p></w:txbxContent>');
	});
});
