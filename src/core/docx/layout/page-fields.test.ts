import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
	loadDocx,
	parseParagraphStyleCatalog,
	parseRunStyleCatalog,
	signedTwips,
	type DocumentModel,
	type HeaderFooterContent,
} from '../index';
import { letterheadFixture } from '../test-support/letterhead-fixture';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import { createPageStories } from './story-flow';
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
	const second = model.blocks[1];
	if (second?.type !== 'paragraph') throw new Error('Expected a second letterhead paragraph');
	second.pageBreakBefore = true;
	const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 16 };
	const pages = layoutDocumentModel(model, measurer).pages;
	expect(pages).toHaveLength(2);
	const first = pages[0]!;
	// The styleless Word package inherits the Normal template's line and paragraph spacing.
	expect(first.marginTopPx).toBeCloseTo(708 / 15 + 11 * 16 * (278 / 240) + 11 * (160 / 15));
	expect(first.columns[0]?.blocks[0]?.yPx).toBe(0);
	expect(pages[1]?.marginTopPx).toBe(96);

	section.marginTopTwips = signedTwips(-1440);
	expect(layoutDocumentModel(model, measurer).pages[0]?.marginTopPx).toBe(96);
});

it('updates story DATE and TIME fields in their result language at the page time', async () => {
	const { model } = await loadDocx(await letterheadFixture());
	const section = model.sections?.[0];
	if (!section) throw new Error('Expected a section');
	delete section.titlePage;
	const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="German"><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:style>
</w:styles>`;
	model.paragraphStyles = parseParagraphStyleCatalog(styles);
	model.characterStyles = parseRunStyleCatalog(styles);
	const date = 'DATE \\@ "dddd d MMMM yyyy"';
	const time = 'TIME \\@ "dddd HH:mm"';
	section.headers = {
		default: {
			blocks: [
				{
					type: 'paragraph',
					id: 'header',
					style: 'German',
					runs: [
						{ text: '', fieldChar: 'begin', language: 'en-US' },
						{ text: '', fieldCode: date, language: 'en-US' },
						{ text: '', fieldChar: 'separate' },
						{ text: 'Saved', field: { instr: date } },
						{ text: '', fieldChar: 'end', language: 'en-US' },
					],
				},
			],
		},
	};
	section.footers = {
		default: {
			blocks: [
				{
					type: 'paragraph',
					id: 'footer',
					runs: [
						{ text: 'Saved', field: { instr: time, simple: true } },
						{ text: ' / ' },
						{ text: 'Saved', field: { instr: time, simple: true }, fieldFlags: { locked: true } },
					],
				},
			],
		},
	};
	const measurer = createFakeMeasurer();
	const { pages } = layoutDocumentModel(model, measurer);
	const stories = createPageStories(
		model,
		pages,
		measurer,
		{},
		undefined,
		new Date(2026, 10, 26, 15, 5),
	);
	expect(
		stories(0).map(({ blocks }) =>
			blocks
				.flatMap((block) =>
					block.kind === 'paragraph'
						? block.lines.flatMap((line) => line.fragments.map(({ text }) => text))
						: [],
				)
				.join(''),
		),
	).toEqual(['Donnerstag 26 November 2026', 'jeudi 15:05 / Saved']);
});
