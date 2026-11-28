import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { finalReviewStructure } from '../review-final-structure';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFakeMeasurer();
const stamp = 'w:author="Author" w:date="2026-01-01T00:00:00Z"';

async function docx(body: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

/** The strings painted on the pages, in order. */
function shown(bytes: Uint8Array): string[] {
	const raw = Buffer.from(bytes).toString('latin1');
	const out: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		if (
			!/\/Filter \/FlateDecode/.test(match[1]!) ||
			/\/Subtype \/(Image|Form)|\/Length1|\/N 3/.test(match[1]!)
		)
			continue;
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const content = inflateSync(bytes.subarray(start, start + length)).toString('latin1');
		for (const text of content.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)) out.push(text[1]!);
	}
	return out;
}
const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const deletedMark = (id: number, content: string) =>
	`<w:p><w:pPr><w:rPr><w:del w:id="${id}" ${stamp}/></w:rPr></w:pPr>${content}</w:p>`;

describe('exportDocxPdf without markup: deleted structure', () => {
	it('joins a paragraph whose mark is deleted to the next one', async () => {
		const model = await docx(
			p('First.') +
				deletedMark(
					1,
					`<w:del w:id="2" ${stamp}><w:r><w:delText>gone text</w:delText></w:r></w:del>`,
				) +
				p('Third.') +
				deletedMark(3, '<w:r><w:t xml:space="preserve">Joined, </w:t></w:r>') +
				p('so one line.'),
		);
		const final = finalReviewStructure(model);
		expect(final.blocks).toHaveLength(3);
		// The source model is not changed.
		expect(model.blocks).toHaveLength(5);
		const text = shown((await exportDocxPdf(model, { measurer })).bytes).join('');
		expect(text).not.toContain('gone');
		expect(text).toContain('Third');
		expect(text).toContain('Joined');
		const marked = shown((await exportDocxPdf(model, { measurer, markup: 'all' })).bytes).join('');
		expect(marked).toContain('gone');
	});
});
