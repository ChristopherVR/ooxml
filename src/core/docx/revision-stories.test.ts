import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { acceptAllRevisions, listRevisions, rejectAllRevisions } from './revision-commands';
import { documentBlockLists } from './document-paragraphs';
import type { DocumentModel, Paragraph } from './model';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const storyNames = ['body', 'header', 'footer', 'footnote', 'endnote'] as const;

function content(name: string): string {
	return `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="${name}-p" w:author="Ada"><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="${name}-r" w:author="Ada"><w:rPr/></w:rPrChange></w:rPr><w:t>${name}</w:t></w:r><w:ins w:id="${name}-i" w:author="Grace"><w:r><w:t>added</w:t></w:r></w:ins></w:p>`;
}

async function fixture() {
	const zip = new JSZip();
	const refs =
		'<w:headerReference w:type="default" r:id="h"/><w:footerReference w:type="default" r:id="f"/>';
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>${content('body')}<w:p><w:pPr><w:sectPr>${refs}</w:sectPr></w:pPr></w:p><w:sectPr>${refs}</w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h" Type="${r}/header" Target="header1.xml"/><Relationship Id="f" Type="${r}/footer" Target="footer1.xml"/></Relationships>`,
	);
	zip.file('word/header1.xml', `<w:hdr xmlns:w="${w}">${content('header')}</w:hdr>`);
	zip.file('word/footer1.xml', `<w:ftr xmlns:w="${w}">${content('footer')}</w:ftr>`);
	for (const kind of ['footnote', 'endnote'])
		zip.file(
			`word/${kind}s.xml`,
			`<w:${kind}s xmlns:w="${w}"><w:${kind} w:id="2">${content(kind)}</w:${kind}></w:${kind}s>`,
		);
	return zip.generateAsync({ type: 'uint8array' });
}

function paragraphs(model: DocumentModel): Paragraph[] {
	return documentBlockLists(model).flatMap((blocks) =>
		blocks.flatMap((block) =>
			block.type === 'paragraph'
				? [block]
				: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
		),
	);
}

describe('review commands across document stories', () => {
	it.each(['accept', 'reject'] as const)(
		'resolves and exports %s for shared headers, footers and notes',
		async (mode) => {
			const loaded = await loadDocx(await fixture());
			const before = structuredClone(loaded.model);
			for (const name of storyNames)
				expect(listRevisions(loaded.model).some((entry) => entry.id === `${name}-p`)).toBe(true);
			const resolved = (mode === 'accept' ? acceptAllRevisions : rejectAllRevisions)(loaded.model);
			expect(loaded.model).toEqual(before);
			expect(listRevisions(resolved)).toEqual([]);
			const reopened = await loadDocx(await loaded.save(resolved));
			expect(listRevisions(reopened.model)).toEqual([]);
			for (const name of storyNames) {
				const matches = paragraphs(reopened.model).filter(
					(paragraph) => paragraph.runs[0]?.text === name,
				);
				expect(matches.length).toBeGreaterThan(0);
				for (const paragraph of matches) {
					expect(paragraph.align).toBe(mode === 'accept' ? 'center' : undefined);
					expect(paragraph.runs[0]?.bold).toBe(mode === 'accept' ? true : undefined);
					expect(paragraph.runs.map((run) => run.text).join('')).toBe(
						name + (mode === 'accept' ? 'added' : ''),
					);
				}
			}
		},
	);

	it('merges revised paragraph marks within each story and table cell', async () => {
		const model = (await loadDocx(await fixture())).model;
		for (const blocks of documentBlockLists(model)) {
			blocks.splice(
				0,
				blocks.length,
				{
					type: 'paragraph',
					id: `${blocks[0]?.id}-first`,
					runs: [{ text: 'One' }],
					markRevision: { id: `${blocks[0]?.id}-mark`, kind: 'insert', author: 'Ada' },
				},
				{ type: 'paragraph', id: `${blocks[0]?.id}-second`, runs: [{ text: 'Two' }] },
			);
		}
		model.blocks.push({
			type: 'table',
			id: 'table',
			rows: [
				[
					{
						paragraphs: [
							{
								type: 'paragraph',
								id: 'cell1',
								runs: [{ text: 'Cell' }],
								markRevision: { id: 'cell-mark', kind: 'insert', author: 'Ada' },
							},
							{ type: 'paragraph', id: 'cell2', runs: [{ text: 'Next' }] },
						],
					},
				],
			],
		});
		const resolved = rejectAllRevisions(model);
		expect(listRevisions(resolved)).toEqual([]);
		expect(
			paragraphs(resolved).map((paragraph) => paragraph.runs.map((run) => run.text).join('')),
		).toEqual(['OneTwo', 'CellNext', 'OneTwo', 'OneTwo', 'OneTwo', 'OneTwo', 'OneTwo', 'OneTwo']);
	});
});
