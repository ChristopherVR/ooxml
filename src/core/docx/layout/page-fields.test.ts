import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx, signedTwips, type DocumentModel, type HeaderFooterContent } from '../index';
import { letterheadFixture } from '../test-support/letterhead-fixture';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import type { LayoutPageBox } from './result';
import {
	headerFooterForPage,
	headerFooterSlotForPage,
	pageNumbers,
	pageNumberValues,
	sectionPageCounts,
	fieldDisplayText,
} from './page-fields';

const fixture = (name: string) =>
	new URL(`./fixtures/continuous-page-fields/${name}`, import.meta.url);
const evidence = JSON.parse(await readFile(fixture('evidence.json'), 'utf8')) as {
	cases: { name: string; pages: { header: string; footer: string }[] }[];
};
function text(
	content: HeaderFooterContent | undefined,
	page: string,
	sectionPages: number,
	numPages: number,
): string {
	return content!.blocks
		.map((block) =>
			block.type === 'paragraph'
				? block.runs
						.map((run) => {
							return fieldDisplayText(run, {
								page,
								sectionPages: String(sectionPages),
								numPages: String(numPages),
							});
						})
						.join('')
				: '',
		)
		.join('');
}

describe('native Word page-field references', () => {
	it.each(evidence.cases)('matches visible headers and page fields in $name', async (reference) => {
		const loaded = await loadDocx(
			new Uint8Array(await readFile(fixture(`${reference.name}.docx`))),
		);
		const result = layoutDocumentModel(loaded.model, createFakeMeasurer());
		const labels = pageNumbers(loaded.model, result.pages);
		const values = pageNumberValues(loaded.model, result.pages);
		const counts = sectionPageCounts(result.pages);
		expect(result.pages).toHaveLength(reference.pages.length);
		const actual = result.pages.map((page, index) => ({
			header: text(
				headerFooterForPage(loaded.model, page, values[index]!, 'headers'),
				labels[index]!,
				counts.get(page.sectionIndex)!,
				result.pages.length,
			),
			footer: text(
				headerFooterForPage(loaded.model, page, values[index]!, 'footers'),
				labels[index]!,
				counts.get(page.sectionIndex)!,
				result.pages.length,
			),
		}));
		expect(actual).toEqual(reference.pages);
	});
});

describe('headerFooterSlotForPage', () => {
	const page = (pageInSection: number) => ({ sectionIndex: 0, pageInSection }) as LayoutPageBox;
	it('picks first, even and default slots', () => {
		const model = { sections: [{ titlePage: true }], evenAndOddHeaders: true } as DocumentModel;
		expect(headerFooterSlotForPage(model, page(0), 1)).toBe('first');
		expect(headerFooterSlotForPage(model, page(1), 2)).toBe('even');
		expect(headerFooterSlotForPage(model, page(2), 3)).toBe('default');
		expect(headerFooterSlotForPage({ sections: [{}] } as DocumentModel, page(1), 2)).toBe(
			'default',
		);
	});
});

it('clears the first-page header paragraphs without counting the floating letterhead shapes', async () => {
	const { model } = await loadDocx(await letterheadFixture());
	const section = model.sections![0]!;
	const header = section.headers!.first!;
	header.blocks.unshift(
		...Array.from({ length: 10 }, (_, index) => ({
			type: 'paragraph' as const,
			id: `blank-${index}`,
			runs: [],
		})),
	);
	const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 16 };
	const first = layoutDocumentModel(model, measurer).pages[0]!;
	expect(first.marginTopPx).toBeCloseTo(708 / 15 + 11 * 16);
	expect(first.columns[0]?.blocks[0]?.yPx).toBe(0);

	section.marginTopTwips = signedTwips(-1440);
	expect(layoutDocumentModel(model, measurer).pages[0]?.marginTopPx).toBe(96);
});
