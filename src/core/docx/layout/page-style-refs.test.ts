import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { parseStyleReference } from '../style-ref';
import { at, expectParagraph, must } from '../test-support/access';
import { complexField, paragraph, styleRefFixture } from '../test-support/style-ref-fixture';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import { pageStyleRefs } from './page-style-refs';

const BREAK = '<w:r><w:br w:type="page"/></w:r>';
const PAGE_BREAK = `<w:p>${BREAK}</w:p>`;
const styled = (text: string) =>
	`<w:r><w:rPr><w:rStyle w:val="EntryName"/></w:rPr><w:t>${text}</w:t></w:r>`;
const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const entry = (name: string, rest: string) =>
	paragraph('', undefined, `${styled(name)}${plain(rest)}`);

async function layout(body: string, header?: string) {
	const { model } = await loadDocx(await styleRefFixture({ body, ...(header ? { header } : {}) }));
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	return { model, pages, refs: pageStyleRefs(model, pages) };
}
const reference = (instr: string) => parseStyleReference(instr)!;
const each = (refs: ReturnType<typeof pageStyleRefs>, pages: number, instr: string) =>
	Array.from({ length: pages }, (_, page) => refs.result(page, reference(instr)));

describe('STYLEREF in headers and footers, page by page', () => {
	it('takes the first text on the page, with \\l the last, else the nearest before the page', async () => {
		const { pages, refs } = await layout(
			paragraph('Alpha', 'Heading1') +
				paragraph('Text') +
				PAGE_BREAK +
				paragraph('Second page') +
				PAGE_BREAK +
				paragraph('Beta', 'Heading1') +
				paragraph('Gamma', 'Heading1'),
		);
		expect(pages).toHaveLength(3);
		expect(each(refs, 3, 'STYLEREF "heading 1"')).toEqual(['Alpha', 'Alpha', 'Beta']);
		expect(each(refs, 3, 'STYLEREF 1 \\l')).toEqual(['Alpha', 'Alpha', 'Gamma']);
	});

	it('takes the nearest text after the page when none is on or before it', async () => {
		const { pages, refs } = await layout(
			paragraph('Cover') + PAGE_BREAK + paragraph('DOC-7', 'DocNumber'),
		);
		expect(pages).toHaveLength(2);
		expect(each(refs, 2, 'STYLEREF "Doc Number"')).toEqual(['DOC-7', 'DOC-7']);
		expect(each(refs, 2, 'STYLEREF Missing')).toEqual([undefined, undefined]);
	});

	it('reads stretches of a character style: the first on the page, or the last with \\l', async () => {
		const { pages, refs } = await layout(
			entry('Ash', ', a tree. ') +
				entry('Aspen', ', another.') +
				PAGE_BREAK +
				paragraph('Pictures only') +
				PAGE_BREAK +
				entry('Birch', '.'),
		);
		expect(pages).toHaveLength(3);
		expect(each(refs, 3, 'STYLEREF "Entry Name"')).toEqual(['Ash', 'Aspen', 'Birch']);
		expect(each(refs, 3, 'STYLEREF "Entry Name" \\l')).toEqual(['Aspen', 'Aspen', 'Birch']);
	});

	it('reads only the stretches of a character style that the page shows of a paragraph', async () => {
		const { pages, refs } = await layout(
			paragraph(
				'',
				undefined,
				`${styled('Ash')}${plain(', a tree. ')}${BREAK}${styled('Aspen')}${plain(', another.')}`,
			) + entry('Birch', '.'),
		);
		expect(pages).toHaveLength(2);
		expect(each(refs, 2, 'STYLEREF "Entry Name"')).toEqual(['Ash', 'Aspen']);
		expect(each(refs, 2, 'STYLEREF "Entry Name" \\l')).toEqual(['Ash', 'Birch']);
	});

	it('looks before and after a page through the parts of a paragraph on other pages', async () => {
		const { pages, refs } = await layout(
			paragraph(
				'',
				undefined,
				`${styled('Ash')}${plain(' grows.')}${BREAK}${plain('Its bark.')}${BREAK}${styled('Elm')}`,
			) +
				PAGE_BREAK +
				paragraph('', undefined, `${plain('Roots.')}${BREAK}${styled('Oak')}`),
		);
		expect(pages).toHaveLength(5);
		expect(each(refs, 5, 'STYLEREF "Entry Name"')).toEqual(['Ash', 'Ash', 'Elm', 'Elm', 'Oak']);
		expect(each(refs, 5, 'STYLEREF "Entry Name" \\l')).toEqual(['Ash', 'Ash', 'Elm', 'Elm', 'Oak']);
	});

	it('passes over a paragraph in the style that has no text', async () => {
		const { pages, refs } = await layout(
			paragraph('DOC-7', 'DocNumber') + PAGE_BREAK + paragraph('', 'DocNumber') + paragraph('Text'),
		);
		expect(pages).toHaveLength(2);
		expect(each(refs, 2, 'STYLEREF "Doc Number"')).toEqual(['DOC-7', 'DOC-7']);
	});

	it('fills a header with the results of the page it is drawn on', async () => {
		const header = paragraph(
			'',
			undefined,
			complexField('STYLEREF 1 \\* MERGEFORMAT', ['Old', ' title']),
		);
		const { model, refs } = await layout(
			paragraph('Alpha', 'Heading1') + PAGE_BREAK + paragraph('Beta', 'Heading1'),
			header,
		);
		const content = must(at(model.sections, 0).headers?.default, 'header');
		const text = (page: number) =>
			expectParagraph(refs.content(content, page).blocks[0])
				.runs.map((run) => run.text)
				.join('');
		expect([text(0), text(1)]).toEqual(['Alpha', 'Beta']);
		// The model's own header keeps its saved result.
		expect(
			expectParagraph(content.blocks[0])
				.runs.map((run) => run.text)
				.join(''),
		).toBe('Old title');
		const locked = structuredClone(content);
		expectParagraph(locked.blocks[0]).runs[0]!.fieldFlags = { locked: true };
		expect(refs.content(locked, 1)).toBe(locked);
	});
});
