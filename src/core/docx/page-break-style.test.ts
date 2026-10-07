import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	createDocument,
	loadDocx,
	parseParagraphStyleCatalog,
	resolveParagraphFormatting,
	saveDocx,
} from './index';
import type { Paragraph } from './model';
import { layoutDocumentModel } from './layout/layout';
import { createFakeMeasurer } from './layout/measure';
import { paragraphAttrs, paragraphFromAttrs } from './ui/paragraph-attributes';
import { parseDirectParagraphProperties } from './paragraph-properties';
import { parseXml, WORD_NS } from './xml';

const native = () =>
	readFile(new URL('./__fixtures__/page-break-style/page-break-style.docx', import.meta.url));
const pageBlocks = (model: Awaited<ReturnType<typeof loadDocx>>['model']) =>
	layoutDocumentModel(model, createFakeMeasurer()).pages.map((page) =>
		page.columns.flatMap((column) => column.blocks.map((block) => block.blockId)),
	);

describe('style-inherited page breaks', () => {
	it('matches native page 1, page 2, page 2 and retains explicit off through editor conversion', async () => {
		const loaded = await loadDocx(new Uint8Array(await native()));
		const paragraphs = loaded.model.blocks as Paragraph[];
		expect(paragraphs.map((paragraph) => paragraph.pageBreakBefore)).toEqual([
			undefined,
			undefined,
			false,
		]);
		expect(
			paragraphs.map(
				(paragraph) =>
					resolveParagraphFormatting(paragraph, loaded.model.paragraphStyles!).pageBreakBefore,
			),
		).toEqual([undefined, true, false]);
		expect(pageBlocks(loaded.model)).toEqual([['p0'], ['p1', 'p2']]);
		for (const paragraph of paragraphs)
			expect(paragraphFromAttrs(paragraphAttrs(paragraph), paragraph.id, paragraph.runs)).toEqual(
				paragraph,
			);
	});
	it('preserves source style XML and pagination after text edits and package export', async () => {
		const bytes = new Uint8Array(await native());
		const loaded = await loadDocx(bytes);
		const originalStyles = await (
			await JSZip.loadAsync(bytes)
		)
			.file('word/styles.xml')!
			.async('string');
		const paragraph = loaded.model.blocks[2] as Paragraph;
		paragraph.runs[0]!.text += '!';
		const exported = await loaded.save(loaded.model);
		expect(await (await JSZip.loadAsync(exported)).file('word/styles.xml')!.async('string')).toBe(
			originalStyles,
		);
		const reopened = await loadDocx(exported);
		expect(pageBlocks(reopened.model)).toEqual([['p0'], ['p1', 'p2']]);
		expect(reopened.model.blocks[2]).toMatchObject({ pageBreakBefore: false });
	});
	it('writes explicit off for a new paragraph and keeps absence distinct', async () => {
		const model = createDocument();
		model.blocks = [
			{ type: 'paragraph', id: 'off', runs: [{ text: 'Off' }], pageBreakBefore: false },
			{ type: 'paragraph', id: 'inherit', runs: [{ text: 'Inherited' }] },
		];
		const reopened = await loadDocx(await saveDocx(model));
		expect((reopened.model.blocks[0] as Paragraph).pageBreakBefore).toBe(false);
		expect((reopened.model.blocks[1] as Paragraph).pageBreakBefore).toBeUndefined();
	});
	it.each(['0', 'false', 'off'])(
		'reads explicit %s and resolves it over a style chain',
		(value) => {
			const direct = parseDirectParagraphProperties(
				parseXml(`<w:pPr xmlns:w="${WORD_NS}"><w:pageBreakBefore w:val="${value}"/></w:pPr>`)
					.documentElement,
			);
			const catalog = parseParagraphStyleCatalog(
				`<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:pageBreakBefore/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/></w:style></w:styles>`,
			);
			expect(
				resolveParagraphFormatting(
					{ type: 'paragraph', id: 'p', runs: [], style: 'Derived' },
					catalog,
				).pageBreakBefore,
			).toBe(true);
			expect(
				resolveParagraphFormatting(
					{ type: 'paragraph', id: 'p', runs: [], style: 'Derived', ...direct },
					catalog,
				).pageBreakBefore,
			).toBe(false);
		},
	);
});
