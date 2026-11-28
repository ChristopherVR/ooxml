// Every nesting level of a tags-appearance content control gets its own start and end tag
// (issue 374), for controls inside a paragraph and for controls around paragraphs.
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (name: string, id: number, content: string, tags = true) =>
	`<w:sdt><w:sdtPr><w:alias w:val="${name}"/><w:tag w:val="${name}"/><w:id w:val="${id}"/>` +
	(tags ? `<w15:appearance w15:val="tags"/>` : '') +
	`</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

async function fragmentsOf(body: string): Promise<string[][]> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:w15="${W15}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	return pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.map((block) =>
				block.kind !== 'paragraph'
					? []
					: block.lines.flatMap((line) =>
							line.fragments.map((fragment) =>
								fragment.contentControlTag
									? `${fragment.contentControlTag}:${fragment.text}`
									: fragment.text,
							),
						),
			),
		),
	);
}
const tags = (fragments: string[]) => fragments.filter((text) => /^(start|end):/.test(text));
const indexOfText = (fragments: string[], word: string) =>
	fragments.findIndex((part) => part.includes(word));

it('tags every level of inline controls nested in a paragraph', async () => {
	const [first, second] = await fragmentsOf(
		`<w:p>${run('Lorem ')}${sdt('Outer', 101, sdt('Inner', 100, run('ipsum dolor')) + run(' sit amet'))}${run(' consectetur.')}</w:p>` +
			`<w:p>${run('Lorem ')}${sdt('Tagged', 103, run('ipsum ') + sdt('Plain', 102, run('dolor sit'), false))}${run(' amet.')}</w:p>`,
	);
	expect(tags(first!)).toEqual(['start:Outer', 'start:Inner', 'end:', 'end:']);
	// The inner tag closes before the text the outer control holds after it.
	const closes = first!.map((part, index) => (part === 'end:' ? index : -1)).filter((i) => i >= 0);
	expect(closes[0]!).toBeLessThan(indexOfText(first!, 'sit'));
	expect(closes[1]!).toBeGreaterThan(indexOfText(first!, 'sit'));
	// A control without tags inside a tagged one does not end it.
	expect(tags(second!)).toEqual(['start:Tagged', 'end:']);
	expect(second!.indexOf('end:')).toBeGreaterThan(indexOfText(second!, 'sit'));
});

it('tags every level of block controls around paragraphs', async () => {
	const paragraphs = await fragmentsOf(
		sdt('Outer', 1, sdt('Inner', 2, `<w:p>${run('One')}</w:p>`) + `<w:p>${run('Two')}</w:p>`),
	);
	expect(paragraphs.map(tags)).toEqual([
		['start:Outer', 'start:Inner', 'end:Inner'],
		['end:Outer'],
	]);
});
