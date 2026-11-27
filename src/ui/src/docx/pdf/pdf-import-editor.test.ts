// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Block, type Paragraph } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { importPdf } from 'ooxml-core/docx/pdf-import';
import { DocxEditorElement, registerDocxEditor } from '../index';
import { exportDocxPdf } from 'ooxml-core/docx/pdf-export';

registerDocxEditor();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const textOf = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');
const words = (text: string) => text.replace(/\s+/g, ' ').trim();

const BODY = [
	'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>Round trip</w:t></w:r></w:p>',
	...Array.from(
		{ length: 3 },
		(_, index) =>
			`<w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:t xml:space="preserve">Paragraph ${index + 1} has enough words to wrap over more than one line of the page, so the conversion must join its lines back into one paragraph.</w:t></w:r></w:p>`,
	),
	'<w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>',
].join('');

async function source() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

afterEach(() => document.body.replaceChildren());

describe('a PDF from the Word PDF export, opened again', () => {
	it('keeps its text, paragraphs, headings and formatting', async () => {
		const { model } = await source();
		const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer() });
		const converted = await importPdf(bytes);
		const back = paragraphs(converted.model.blocks);
		expect(words(back.map(textOf).join(' '))).toBe(
			words(paragraphs(model.blocks).map(textOf).join(' ')),
		);
		expect(back.map(textOf).filter((text) => text.startsWith('Paragraph'))).toHaveLength(3);
		expect(back[0]).toMatchObject({ style: 'Heading1' });
		const runs = back[back.length - 1]!.runs;
		expect(runs.find((run) => run.text.trim() === 'bold')).toMatchObject({ bold: true });
		expect(runs.find((run) => run.text.trim() === 'italic')).toMatchObject({ italic: true });
	});
});

describe('<docx-editor> opening a PDF', () => {
	it('converts it into an editable document named .docx, with the conversion notice', async () => {
		const { model } = await source();
		const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer() });
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		await editor.load(bytes);
		const opened = editor.documentModel!;
		expect(textOf(paragraphs(opened.blocks)[0]!)).toBe('Round trip');
		expect(opened.warnings[0]).toMatch(/converted from a PDF/);
		const saved = await loadDocx(await editor.saveBytes());
		expect(paragraphs(saved.model.blocks).map(textOf)).toContain('Plain bold and italic.');
	});
});
